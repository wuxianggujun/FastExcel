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
    
    // 优化的字符串缓冲区 - 小字符串优化
    struct OptimizedStringBuffer {
        static constexpr size_t SMALL_SIZE = 64;
        char small_buffer_[SMALL_SIZE];
        std::string large_buffer_;
        size_t size_ = 0;
        bool using_small_ = true;
        
        void clear() {
            size_ = 0;
            using_small_ = true;
            if (!large_buffer_.empty()) {
                large_buffer_.clear();
            }
        }
        
        void append(std::string_view data) {
            if (using_small_ && size_ + data.size() <= SMALL_SIZE) {
                std::memcpy(small_buffer_ + size_, data.data(), data.size());
                size_ += data.size();
            } else {
                // 切换到大缓冲区
                if (using_small_) {
                    large_buffer_.assign(small_buffer_, size_);
                    using_small_ = false;
                }
                large_buffer_.append(data);
                size_ = large_buffer_.size();
            }
        }
        
        std::string_view view() const {
            if (using_small_) {
                return std::string_view(small_buffer_, size_);
            } else {
                return large_buffer_;
            }
        }
        
        const char* data() const {
            return using_small_ ? small_buffer_ : large_buffer_.data();
        }
        
        bool empty() const { return size_ == 0; }
    } current_cell_value_;
    
    // 统计信息
    size_t cells_processed_ = 0;
    
    // 辅助方法
    uint32_t parseColumnReference(std::string_view cell_ref);
    void processCellValue();
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
    }
};

}} // namespace fastexcel::reader