#include "fastexcel/read/ReadWorksheet.hpp"
#include "fastexcel/core/Logger.hpp"
#include <sstream>
#include <algorithm>

namespace fastexcel::read {

// 构造函数
ReadWorksheet::ReadWorksheet(reader::XLSXReader* reader, const std::string& name)
    : reader_(reader)
    , name_(name)
    , row_count_(0)
    , column_count_(0)
    , cell_cache_(1024)  // 1024个槽位的缓存
    , cache_hits_(0)
    , cache_misses_(0) {
    
    if (!reader_) {
        throw std::invalid_argument("Reader cannot be null");
    }
    
    // 获取工作表
    worksheet_ = reader_->getWorksheet(name);
    if (!worksheet_) {
        throw std::runtime_error("Failed to load worksheet: " + name);
    }
    
    // 计算维度
    calculateDimensions();
    
    LOG_DEBUG("ReadWorksheet created: {}, dimensions: {}x{}", 
             name, row_count_, column_count_);
}

// 析构函数
ReadWorksheet::~ReadWorksheet() {
    // 输出缓存统计
    if (cache_hits_ + cache_misses_ > 0) {
        double hit_rate = static_cast<double>(cache_hits_) / 
                         (cache_hits_ + cache_misses_) * 100.0;
        LOG_TRACE("Worksheet {} cache - Hits: {}, Misses: {}, Hit rate: {:.2f}%",
                 name_, cache_hits_, cache_misses_, hit_rate);
    }
}

// 获取工作表名称
std::string ReadWorksheet::getName() const {
    return name_;
}

// 获取行数
size_t ReadWorksheet::getRowCount() const {
    return row_count_;
}

// 获取列数
size_t ReadWorksheet::getColumnCount() const {
    return column_count_;
}

// 读取字符串值
std::string ReadWorksheet::readString(size_t row, size_t col) const {
    // 边界检查
    if (row >= row_count_ || col >= column_count_) {
        LOG_TRACE("Cell out of bounds: ({}, {})", row, col);
        return "";
    }
    
    // 尝试从缓存获取
    auto cached = getFromCache(row, col);
    if (cached.has_value()) {
        return std::get<std::string>(cached.value());
    }
    
    // 从工作表读取
    auto cell = worksheet_->getCell(row, col);
    if (!cell) {
        return "";
    }
    
    std::string value = cell->getString();
    
    // 添加到缓存
    addToCache(row, col, value);
    
    return value;
}

// 读取数值
double ReadWorksheet::readNumber(size_t row, size_t col) const {
    // 边界检查
    if (row >= row_count_ || col >= column_count_) {
        LOG_TRACE("Cell out of bounds: ({}, {})", row, col);
        return 0.0;
    }
    
    // 尝试从缓存获取
    auto cached = getFromCache(row, col);
    if (cached.has_value()) {
        if (std::holds_alternative<double>(cached.value())) {
            return std::get<double>(cached.value());
        }
        // 尝试转换字符串为数字
        try {
            return std::stod(std::get<std::string>(cached.value()));
        } catch (...) {
            return 0.0;
        }
    }
    
    // 从工作表读取
    auto cell = worksheet_->getCell(row, col);
    if (!cell) {
        return 0.0;
    }
    
    double value = cell->getNumber();
    
    // 添加到缓存
    addToCache(row, col, value);
    
    return value;
}

// 读取布尔值
bool ReadWorksheet::readBool(size_t row, size_t col) const {
    // 边界检查
    if (row >= row_count_ || col >= column_count_) {
        LOG_TRACE("Cell out of bounds: ({}, {})", row, col);
        return false;
    }
    
    // 从工作表读取
    auto cell = worksheet_->getCell(row, col);
    if (!cell) {
        return false;
    }
    
    return cell->getBool();
}

// 读取日期时间
std::chrono::system_clock::time_point ReadWorksheet::readDateTime(size_t row, size_t col) const {
    // 边界检查
    if (row >= row_count_ || col >= column_count_) {
        LOG_TRACE("Cell out of bounds: ({}, {})", row, col);
        return std::chrono::system_clock::time_point{};
    }
    
    // 从工作表读取
    auto cell = worksheet_->getCell(row, col);
    if (!cell) {
        return std::chrono::system_clock::time_point{};
    }
    
    // Excel日期是从1900年1月1日开始的天数
    double excel_date = cell->getNumber();
    
    // 转换为系统时间点
    // Excel epoch: 1900-01-01 (实际上是1899-12-30因为Excel的bug)
    static const auto excel_epoch = std::chrono::system_clock::from_time_t(-2209161600);
    auto days = std::chrono::duration<double, std::ratio<86400>>(excel_date);
    
    return excel_epoch + std::chrono::duration_cast<std::chrono::system_clock::duration>(days);
}

// 读取公式
std::string ReadWorksheet::readFormula(size_t row, size_t col) const {
    // 边界检查
    if (row >= row_count_ || col >= column_count_) {
        LOG_TRACE("Cell out of bounds: ({}, {})", row, col);
        return "";
    }
    
    // 从工作表读取
    auto cell = worksheet_->getCell(row, col);
    if (!cell) {
        return "";
    }
    
    return cell->getFormula();
}

// 获取单元格类型
CellType ReadWorksheet::getCellType(size_t row, size_t col) const {
    // 边界检查
    if (row >= row_count_ || col >= column_count_) {
        return CellType::EMPTY;
    }
    
    // 从工作表读取
    auto cell = worksheet_->getCell(row, col);
    if (!cell) {
        return CellType::EMPTY;
    }
    
    switch (cell->getType()) {
        case core::CellType::String:
            return CellType::STRING;
        case core::CellType::Number:
            return CellType::NUMBER;
        case core::CellType::Boolean:
            return CellType::BOOLEAN;
        case core::CellType::Date:
            return CellType::DATE;
        case core::CellType::Formula:
            return CellType::FORMULA;
        default:
            return CellType::EMPTY;
    }
}

// 检查单元格是否为空
bool ReadWorksheet::isEmpty(size_t row, size_t col) const {
    return getCellType(row, col) == CellType::EMPTY;
}

// 读取整行
std::vector<std::variant<std::string, double, bool>> ReadWorksheet::readRow(size_t row) const {
    std::vector<std::variant<std::string, double, bool>> result;
    
    if (row >= row_count_) {
        return result;
    }
    
    result.reserve(column_count_);
    
    for (size_t col = 0; col < column_count_; ++col) {
        auto type = getCellType(row, col);
        switch (type) {
            case CellType::STRING:
                result.emplace_back(readString(row, col));
                break;
            case CellType::NUMBER:
            case CellType::DATE:
                result.emplace_back(readNumber(row, col));
                break;
            case CellType::BOOLEAN:
                result.emplace_back(readBool(row, col));
                break;
            default:
                result.emplace_back("");  // 空单元格用空字符串表示
                break;
        }
    }
    
    return result;
}

// 读取整列
std::vector<std::variant<std::string, double, bool>> ReadWorksheet::readColumn(size_t col) const {
    std::vector<std::variant<std::string, double, bool>> result;
    
    if (col >= column_count_) {
        return result;
    }
    
    result.reserve(row_count_);
    
    for (size_t row = 0; row < row_count_; ++row) {
        auto type = getCellType(row, col);
        switch (type) {
            case CellType::STRING:
                result.emplace_back(readString(row, col));
                break;
            case CellType::NUMBER:
            case CellType::DATE:
                result.emplace_back(readNumber(row, col));
                break;
            case CellType::BOOLEAN:
                result.emplace_back(readBool(row, col));
                break;
            default:
                result.emplace_back("");  // 空单元格用空字符串表示
                break;
        }
    }
    
    return result;
}

// 读取范围
std::vector<std::vector<std::variant<std::string, double, bool>>> 
ReadWorksheet::readRange(size_t start_row, size_t start_col, 
                         size_t end_row, size_t end_col) const {
    
    std::vector<std::vector<std::variant<std::string, double, bool>>> result;
    
    // 调整边界
    end_row = std::min(end_row, row_count_ - 1);
    end_col = std::min(end_col, column_count_ - 1);
    
    if (start_row > end_row || start_col > end_col) {
        return result;
    }
    
    result.reserve(end_row - start_row + 1);
    
    for (size_t row = start_row; row <= end_row; ++row) {
        std::vector<std::variant<std::string, double, bool>> row_data;
        row_data.reserve(end_col - start_col + 1);
        
        for (size_t col = start_col; col <= end_col; ++col) {
            auto type = getCellType(row, col);
            switch (type) {
                case CellType::STRING:
                    row_data.emplace_back(readString(row, col));
                    break;
                case CellType::NUMBER:
                case CellType::DATE:
                    row_data.emplace_back(readNumber(row, col));
                    break;
                case CellType::BOOLEAN:
                    row_data.emplace_back(readBool(row, col));
                    break;
                default:
                    row_data.emplace_back("");
                    break;
            }
        }
        
        result.push_back(std::move(row_data));
    }
    
    return result;
}

// 创建行迭代器
std::unique_ptr<RowIterator> ReadWorksheet::createRowIterator() const {
    return std::make_unique<RowIteratorImpl>(this);
}

// 克隆工作表对象
std::unique_ptr<IReadOnlyWorksheet> ReadWorksheet::clone() const {
    return std::make_unique<ReadWorksheet>(reader_, name_);
}

// 计算工作表维度
void ReadWorksheet::calculateDimensions() {
    if (!worksheet_) {
        row_count_ = 0;
        column_count_ = 0;
        return;
    }
    
    // 获取工作表的维度
    auto dimension = worksheet_->getDimension();
    row_count_ = dimension.rows;
    column_count_ = dimension.columns;
    
    // 如果维度信息不可用，扫描工作表
    if (row_count_ == 0 || column_count_ == 0) {
        size_t max_row = 0;
        size_t max_col = 0;
        
        // 扫描所有非空单元格
        for (const auto& [ref, cell] : worksheet_->getCells()) {
            auto [row, col] = core::CellReference::parse(ref);
            max_row = std::max(max_row, row);
            max_col = std::max(max_col, col);
        }
        
        row_count_ = max_row + 1;
        column_count_ = max_col + 1;
    }
}

// 从缓存获取值
std::optional<std::variant<std::string, double, bool>> 
ReadWorksheet::getFromCache(size_t row, size_t col) const {
    size_t hash = hashCell(row, col);
    size_t slot = hash % cell_cache_.size();
    
    auto& entry = cell_cache_[slot];
    if (entry.row == row && entry.col == col && entry.valid) {
        cache_hits_++;
        return entry.value;
    }
    
    cache_misses_++;
    return std::nullopt;
}

// 添加到缓存
void ReadWorksheet::addToCache(size_t row, size_t col, 
                               const std::variant<std::string, double, bool>& value) const {
    size_t hash = hashCell(row, col);
    size_t slot = hash % cell_cache_.size();
    
    cell_cache_[slot] = {row, col, value, true};
}

// 计算单元格哈希值
size_t ReadWorksheet::hashCell(size_t row, size_t col) const {
    // 简单的哈希函数：组合行列索引
    return (row << 16) ^ col;
}

// RowIteratorImpl 实现
class RowIteratorImpl : public RowIterator {
public:
    explicit RowIteratorImpl(const ReadWorksheet* worksheet)
        : worksheet_(worksheet)
        , current_row_(0)
        , total_rows_(worksheet->getRowCount()) {
    }
    
    bool hasNext() const override {
        return current_row_ < total_rows_;
    }
    
    std::vector<std::variant<std::string, double, bool>> next() override {
        if (!hasNext()) {
            throw std::out_of_range("No more rows");
        }
        
        auto row = worksheet_->readRow(current_row_);
        current_row_++;
        return row;
    }
    
    void reset() override {
        current_row_ = 0;
    }
    
    size_t getCurrentRow() const override {
        return current_row_;
    }
    
private:
    const ReadWorksheet* worksheet_;
    size_t current_row_;
    size_t total_rows_;
};

} // namespace fastexcel::read