#include "WorksheetChain.hpp"
#include "Worksheet.hpp"

namespace fastexcel {
namespace core {

// 模板方法的显式实例化和实现
template<typename T>
WorksheetChain& WorksheetChain::setValue(int row, int col, const T& value) {
    worksheet_.setValue<T>(row, col, value);
    return *this;
}

template<typename T>
WorksheetChain& WorksheetChain::setValue(const std::string& address, const T& value) {
    worksheet_.setValue<T>(address, value);
    return *this;
}

template<typename T>
WorksheetChain& WorksheetChain::setRange(int start_row, int start_col, const std::vector<std::vector<T>>& data) {
    worksheet_.setRange<T>(start_row, start_col, data);
    return *this;
}

template<typename T>
WorksheetChain& WorksheetChain::setRange(const std::string& range, const std::vector<std::vector<T>>& data) {
    worksheet_.setRange<T>(range, data);
    return *this;
}

// 非模板方法的实现
WorksheetChain& WorksheetChain::setColumnWidth(int col, double width) {
    worksheet_.setColumnWidth(col, width);
    return *this;
}

WorksheetChain& WorksheetChain::setRowHeight(int row, double height) {
    worksheet_.setRowHeight(row, height);
    return *this;
}

WorksheetChain& WorksheetChain::mergeCells(int first_row, int first_col, int last_row, int last_col) {
    worksheet_.mergeCells(first_row, first_col, last_row, last_col);
    return *this;
}

// 显式实例化常用类型的模板
template WorksheetChain& WorksheetChain::setValue<std::string>(int, int, const std::string&);
template WorksheetChain& WorksheetChain::setValue<double>(int, int, const double&);
template WorksheetChain& WorksheetChain::setValue<int>(int, int, const int&);
template WorksheetChain& WorksheetChain::setValue<bool>(int, int, const bool&);

template WorksheetChain& WorksheetChain::setValue<std::string>(const std::string&, const std::string&);
template WorksheetChain& WorksheetChain::setValue<double>(const std::string&, const double&);
template WorksheetChain& WorksheetChain::setValue<int>(const std::string&, const int&);
template WorksheetChain& WorksheetChain::setValue<bool>(const std::string&, const bool&);

template WorksheetChain& WorksheetChain::setRange<std::string>(int, int, const std::vector<std::vector<std::string>>&);
template WorksheetChain& WorksheetChain::setRange<double>(int, int, const std::vector<std::vector<double>>&);
template WorksheetChain& WorksheetChain::setRange<int>(int, int, const std::vector<std::vector<int>>&);

template WorksheetChain& WorksheetChain::setRange<std::string>(const std::string&, const std::vector<std::vector<std::string>>&);
template WorksheetChain& WorksheetChain::setRange<double>(const std::string&, const std::vector<std::vector<double>>&);
template WorksheetChain& WorksheetChain::setRange<int>(const std::string&, const std::vector<std::vector<int>>&);

}} // namespace fastexcel::core