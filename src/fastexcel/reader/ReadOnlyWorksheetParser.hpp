/**
 * @file ReadOnlyWorksheetParser.hpp
 * @brief 只读模式专用工作表XML流式解析器
 */

#pragma once

#include "BaseSAXParser.hpp"
#include "fastexcel/core/ErrorCode.hpp"
#include "fastexcel/core/ColumnarStorageManager.hpp"
#include "fastexcel/core/WorkbookTypes.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"
#include <memory>
#include <string>
#include <unordered_map>

namespace fastexcel {
namespace reader {

/**
 * @brief 专用于只读模式的工作表XML流式解析器
 * 直接将数据解析到ColumnarStorageManager，完全避免Cell对象创建
 */
class ReadOnlyWorksheetParser : public BaseSAXParser {
private:
    std::shared_ptr<core::ColumnarStorageManager> storage_;
    const std::unordered_map<int, std::string_view>* shared_strings_;
    const core::WorkbookOptions* options_;
    
    // 解析状态
    bool in_sheet_data_ = false;
    bool in_row_ = false;
    bool in_cell_ = false;
    bool in_value_ = false;
    
    // 当前解析的数据
    int current_row_ = 0;
    uint32_t current_col_ = 0;
    std::string_view current_cell_type_;
    std::string current_cell_value_;
    
    // 批量处理优化
    struct BatchCellData {
        uint32_t row, col;
        std::string value;
        std::string_view type;
        
        BatchCellData(uint32_t r, uint32_t c, std::string&& v, std::string_view t)
            : row(r), col(c), value(std::move(v)), type(t) {}
    };
    std::vector<BatchCellData> row_batch_;
    static constexpr size_t BATCH_SIZE = 1000; // 每1000个单元格批量处理一次
    
    // 统计信息
    size_t cells_processed_ = 0;
    
    // 辅助方法
    uint32_t parseColumnReference(std::string_view cell_ref);
    void processCellValue();
    void processBatch(); // 批量处理单元格
    void processBatchStandard(); // 标准批量处理
#ifdef FASTEXCEL_HAS_HIGHWAY
    void processBatchSIMD(); // SIMD优化批量处理
#endif
    void flushBatch();   // 强制处理剩余批量数据
    bool shouldSkipCell(uint32_t row, uint32_t col) const;
    
protected:
    // 重写SAX事件处理方法
    void onStartElement(std::string_view name, 
                       span<const xml::XMLAttribute> attributes, int depth) override;
    void onEndElement(std::string_view name, int depth) override;
    void onText(std::string_view data, int depth) override;
    
public:
    ReadOnlyWorksheetParser() = default;
    ~ReadOnlyWorksheetParser() = default;
    
    /**
     * @brief 设置解析参数
     */
    void configure(std::shared_ptr<core::ColumnarStorageManager> storage,
                  const std::unordered_map<int, std::string_view>* shared_strings,
                  const core::WorkbookOptions* options) {
        storage_ = storage;
        shared_strings_ = shared_strings;
        options_ = options;
        
        // 预分配内存优化
        current_cell_value_.reserve(256); // 预分配单元格值缓冲区
    }
    
    /**
     * @brief 获取处理的单元格数量
     */
    size_t getCellsProcessed() const {
        return cells_processed_;
    }
    
    /**
     * @brief 重置解析状态
     */
    void reset() {
        state_.reset();
        in_sheet_data_ = false;
        in_row_ = false;
        in_cell_ = false;
        in_value_ = false;
        current_row_ = 0;
        current_col_ = 0;
        current_cell_type_ = std::string_view{};
        current_cell_value_.clear();
        cells_processed_ = 0;
        row_batch_.clear();
        row_batch_.reserve(BATCH_SIZE); // 预分配批量缓冲区
    }
};

}} // namespace fastexcel::reader