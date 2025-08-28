#include "WorksheetChain.hpp"
#include "Worksheet.hpp"

namespace fastexcel {
namespace core {

// 模板方法的实现
template<typename T>
WorksheetChain& WorksheetChain::setValue(const Address& address, const T& value) {
    worksheet_.setValue<T>(address, value);
    return *this;
}

template<typename T>
WorksheetChain& WorksheetChain::setRange(const CellRange& range, const std::vector<std::vector<T>>& data) {
    worksheet_.setRange<T>(range, data);
    return *this;
}

// 非模板方法的实现
WorksheetChain& WorksheetChain::setColumnWidth(const Address& col, double width) {
    worksheet_.setColumnWidth(col.getCol(), width);
    return *this;
}

WorksheetChain& WorksheetChain::setRowHeight(const Address& row, double height) {
    worksheet_.setRowHeight(row.getRow(), height);
    return *this;
}

WorksheetChain& WorksheetChain::mergeCells(const CellRange& range) {
    worksheet_.mergeCells(range);
    return *this;
}

// 显式实例化常用类型的模板
template WorksheetChain& WorksheetChain::setValue<std::string>(const Address&, const std::string&);
template WorksheetChain& WorksheetChain::setValue<double>(const Address&, const double&);
template WorksheetChain& WorksheetChain::setValue<int>(const Address&, const int&);
template WorksheetChain& WorksheetChain::setValue<bool>(const Address&, const bool&);

template WorksheetChain& WorksheetChain::setRange<std::string>(const CellRange&, const std::vector<std::vector<std::string>>&);
template WorksheetChain& WorksheetChain::setRange<double>(const CellRange&, const std::vector<std::vector<double>>&);
template WorksheetChain& WorksheetChain::setRange<int>(const CellRange&, const std::vector<std::vector<int>>&);

}} // namespace fastexcel::core