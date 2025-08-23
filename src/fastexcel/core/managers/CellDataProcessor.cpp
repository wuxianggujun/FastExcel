#include "CellDataProcessor.hpp"
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/DirtyManager.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>
#include <cctype>

namespace fastexcel {
namespace core {

CellDataProcessor::CellDataProcessor(std::map<std::pair<int, int>, Cell>& cells, 
                                   CellRangeManager& range_manager,
                                   std::shared_ptr<Workbook> parent_workbook,
                                   int sheet_id)
    : cells_(cells), range_manager_(range_manager), parent_workbook_(parent_workbook), sheet_id_(sheet_id) {
}

Cell& CellDataProcessor::getCell(int row, int col) {
    validateCellPosition(row, col);
    updateUsedRange(row, col);
    return cells_[std::make_pair(row, col)];
}

const Cell& CellDataProcessor::getCell(int row, int col) const {
    validateCellPosition(row, col);
    auto it = cells_.find(std::make_pair(row, col));
    if (it == cells_.end()) {
        static Cell empty_cell;
        return empty_cell;
    }
    return it->second;
}

void CellDataProcessor::setFormula(int row, int col, const std::string& formula, double result) {
    validateCellPosition(row, col);
    updateUsedRange(row, col);
    markDirty();
    getCell(row, col).setFormula(formula, result);
}

void CellDataProcessor::setHyperlink(int row, int col, const std::string& url, const std::string& display_text) {
    validateCellPosition(row, col);
    updateUsedRange(row, col);
    
    // 超链接需要特殊的DirtyManager处理，涉及关系文件
    if (parent_workbook_ && parent_workbook_->getDirtyManager() && sheet_id_ >= 0) {
        std::string sheet_path = fmt::format("xl/worksheets/sheet{}.xml", sheet_id_);
        std::string rels_path = fmt::format("xl/worksheets/_rels/sheet{}.xml.rels", sheet_id_);
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
        parent_workbook_->getDirtyManager()->markDirty(rels_path, DirtyManager::DirtyLevel::CONTENT);
    }
    
    auto& cell = getCell(row, col);
    if (!display_text.empty()) {
        cell.setValue(display_text);
    } else {
        cell.setValue(url);
    }
    cell.setHyperlink(url);
}

void CellDataProcessor::clearCell(int row, int col) {
    auto it = cells_.find(std::make_pair(row, col));
    if (it != cells_.end()) {
        markDirty();
        cells_.erase(it);
    }
}

bool CellDataProcessor::hasCellAt(int row, int col) const {
    auto it = cells_.find(std::make_pair(row, col));
    return it != cells_.end() && (!it->second.isEmpty() || it->second.hasFormat());
}

void CellDataProcessor::copyCell(int src_row, int src_col, int dst_row, int dst_col, bool copy_format) {
    validateCellPosition(src_row, src_col);
    validateCellPosition(dst_row, dst_col);
    
    const auto& src_cell = getCell(src_row, src_col);
    if (src_cell.isEmpty()) {
        return;
    }
    
    auto& dst_cell = getCell(dst_row, dst_col);
    
    // 复制值
    if (src_cell.isString()) {
        dst_cell.setValue(src_cell.getValue<std::string>());
    } else if (src_cell.isNumber()) {
        dst_cell.setValue(src_cell.getValue<double>());
    } else if (src_cell.isBoolean()) {
        dst_cell.setValue(src_cell.getValue<bool>());
    } else if (src_cell.isFormula()) {
        dst_cell.setFormula(src_cell.getFormula(), src_cell.getFormulaResult());
    }
    
    // 复制格式
    if (copy_format && src_cell.hasFormat()) {
        dst_cell.setFormat(src_cell.getFormatDescriptor());
    }
    
    // 复制超链接
    if (src_cell.hasHyperlink()) {
        dst_cell.setHyperlink(src_cell.getHyperlink());
    }
    
    updateUsedRange(dst_row, dst_col);
}

void CellDataProcessor::moveCell(int src_row, int src_col, int dst_row, int dst_col) {
    validateCellPosition(src_row, src_col);
    validateCellPosition(dst_row, dst_col);
    
    if (src_row == dst_row && src_col == dst_col) {
        return;
    }
    
    copyCell(src_row, src_col, dst_row, dst_col, true);
    clearCell(src_row, src_col);
}

void CellDataProcessor::copyRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                                 int dst_row, int dst_col, bool copy_format) {
    validateRange(src_first_row, src_first_col, src_last_row, src_last_col);
    
    int rows = src_last_row - src_first_row + 1;
    int cols = src_last_col - src_first_col + 1;
    
    validateCellPosition(dst_row + rows - 1, dst_col + cols - 1);
    
    for (int r = 0; r < rows; ++r) {
        for (int c = 0; c < cols; ++c) {
            copyCell(src_first_row + r, src_first_col + c,
                    dst_row + r, dst_col + c, copy_format);
        }
    }
}

void CellDataProcessor::moveRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                                 int dst_row, int dst_col) {
    copyRange(src_first_row, src_first_col, src_last_row, src_last_col, dst_row, dst_col, true);
    clearRange(src_first_row, src_first_col, src_last_row, src_last_col);
}

void CellDataProcessor::clearRange(int first_row, int first_col, int last_row, int last_col) {
    validateRange(first_row, first_col, last_row, last_col);
    
    for (int row = first_row; row <= last_row; ++row) {
        for (int col = first_col; col <= last_col; ++col) {
            clearCell(row, col);
        }
    }
}

std::vector<std::pair<int, int>> CellDataProcessor::findCells(const std::string& search_text,
                                                             bool match_case,
                                                             bool match_entire_cell) const {
    std::vector<std::pair<int, int>> results;
    
    for (const auto& [pos, cell] : cells_) {
        if (!cell.isString()) {
            continue;
        }
        
        std::string cell_text = cell.getValue<std::string>();
        std::string target_text = cell_text;
        std::string find_text = search_text;
        
        if (!match_case) {
            std::transform(find_text.begin(), find_text.end(), find_text.begin(),
                         [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
            std::transform(target_text.begin(), target_text.end(), target_text.begin(),
                         [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
        }
        
        bool found = false;
        if (match_entire_cell) {
            found = (target_text == find_text);
        } else {
            found = (target_text.find(find_text) != std::string::npos);
        }
        
        if (found) {
            results.push_back(pos);
        }
    }
    
    return results;
}

int CellDataProcessor::findAndReplace(const std::string& find_text, const std::string& replace_text,
                                     bool match_case, bool match_entire_cell) {
    int replace_count = 0;
    
    for (auto& [pos, cell] : cells_) {
        if (!cell.isString()) {
            continue;
        }
        
        std::string cell_text = cell.getValue<std::string>();
        bool replaced = false;
        
        if (match_entire_cell) {
            std::string target_text = cell_text;
            std::string search_text = find_text;
            
            if (!match_case) {
                std::transform(search_text.begin(), search_text.end(), search_text.begin(),
                             [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
                std::transform(target_text.begin(), target_text.end(), target_text.begin(),
                             [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
            }
            
            if (target_text == search_text) {
                cell.setValue(replace_text);
                replaced = true;
            }
        } else {
            // 部分匹配替换
            size_t pos = 0;
            while ((pos = cell_text.find(find_text, pos)) != std::string::npos) {
                cell_text.replace(pos, find_text.length(), replace_text);
                pos += replace_text.length();
                replaced = true;
            }
            
            if (replaced) {
                cell.setValue(cell_text);
            }
        }
        
        if (replaced) {
            replace_count++;
        }
    }
    
    return replace_count;
}

void CellDataProcessor::sortRange(int first_row, int first_col, int last_row, int last_col,
                                 int sort_column, bool ascending, bool has_header) {
    validateRange(first_row, first_col, last_row, last_col);
    
    int data_start_row = has_header ? first_row + 1 : first_row;
    if (data_start_row > last_row) {
        return;
    }
    
    int sort_col = first_col + sort_column;
    if (sort_col > last_col) {
        FASTEXCEL_THROW_PARAM("Sort column is outside the range");
    }
    
    // 收集需要排序的行数据
    std::vector<std::pair<int, std::map<int, Cell>>> rows_data;
    
    for (int row = data_start_row; row <= last_row; ++row) {
        std::map<int, Cell> row_cells;
        for (int col = first_col; col <= last_col; ++col) {
            auto it = cells_.find(std::make_pair(row, col));
            if (it != cells_.end()) {
                row_cells[col] = std::move(it->second);
                cells_.erase(it);
            }
        }
        rows_data.emplace_back(row, std::move(row_cells));
    }
    
    // 排序逻辑（简化版）
    std::sort(rows_data.begin(), rows_data.end(),
        [sort_col, ascending](const auto& a, const auto& b) {
            const auto& a_cells = a.second;
            const auto& b_cells = b.second;
            
            auto a_it = a_cells.find(sort_col);
            auto b_it = b_cells.find(sort_col);
            
            if (a_it == a_cells.end() && b_it == b_cells.end()) {
                return false;
            }
            if (a_it == a_cells.end()) {
                return ascending;
            }
            if (b_it == b_cells.end()) {
                return !ascending;
            }
            
            const Cell& a_cell = a_it->second;
            const Cell& b_cell = b_it->second;
            
            if (a_cell.isNumber() && b_cell.isNumber()) {
                double a_val = a_cell.getValue<double>();
                double b_val = b_cell.getValue<double>();
                return ascending ? (a_val < b_val) : (a_val > b_val);
            } else if (a_cell.isString() && b_cell.isString()) {
                const std::string& a_str = a_cell.getValue<std::string>();
                const std::string& b_str = b_cell.getValue<std::string>();
                return ascending ? (a_str < b_str) : (a_str > b_str);
            }
            
            return false;
        });
    
    // 将排序后的数据放回
    for (size_t i = 0; i < rows_data.size(); ++i) {
        int target_row = data_start_row + static_cast<int>(i);
        const auto& row_cells = rows_data[i].second;
        
        for (const auto& [col, cell] : row_cells) {
            cells_[std::make_pair(target_row, col)] = std::move(const_cast<Cell&>(cell));
            updateUsedRange(target_row, col);
        }
    }
}

int CellDataProcessor::getCellCount() const {
    int count = 0;
    for (const auto& [pos, cell] : cells_) {
        if (!cell.isEmpty()) {
            count++;
        }
    }
    return count;
}

int CellDataProcessor::getCellCountInRow(int row) const {
    int count = 0;
    for (const auto& [pos, cell] : cells_) {
        if (pos.first == row && !cell.isEmpty()) {
            count++;
        }
    }
    return count;
}

int CellDataProcessor::getCellCountInColumn(int col) const {
    int count = 0;
    for (const auto& [pos, cell] : cells_) {
        if (pos.second == col && !cell.isEmpty()) {
            count++;
        }
    }
    return count;
}

std::pair<int, int> CellDataProcessor::getUsedRange() const {
    return range_manager_.getUsedRowRange().first != -1 ? 
           std::make_pair(range_manager_.getUsedRowRange().second, range_manager_.getUsedColRange().second) :
           std::make_pair(-1, -1);
}

std::tuple<int, int, int, int> CellDataProcessor::getUsedRangeFull() const {
    return range_manager_.getUsedRange();
}

void CellDataProcessor::clearRow(int row) {
    auto it = cells_.begin();
    while (it != cells_.end()) {
        if (it->first.first == row) {
            it = cells_.erase(it);
        } else {
            ++it;
        }
    }
    markDirty();
    FASTEXCEL_LOG_DEBUG("Cleared row {}", row);
}

void CellDataProcessor::clearColumn(int col) {
    auto it = cells_.begin();
    while (it != cells_.end()) {
        if (it->first.second == col) {
            it = cells_.erase(it);
        } else {
            ++it;
        }
    }
    markDirty();
    FASTEXCEL_LOG_DEBUG("Cleared column {}", col);
}

void CellDataProcessor::clearAll() {
    cells_.clear();
    markDirty();
    FASTEXCEL_LOG_DEBUG("Cleared all cells");
}

void CellDataProcessor::validateCellPosition(int row, int col) const {
    FASTEXCEL_VALIDATE_CELL_POSITION(row, col);
}

void CellDataProcessor::validateRange(int first_row, int first_col, int last_row, int last_col) const {
    FASTEXCEL_VALIDATE_RANGE(first_row, first_col, last_row, last_col);
}

void CellDataProcessor::updateUsedRange(int row, int col) {
    range_manager_.updateRange(row, col);
}

void CellDataProcessor::markDirty() const {
    if (parent_workbook_ && parent_workbook_->getDirtyManager() && sheet_id_ >= 0) {
        std::string sheet_path = fmt::format("xl/worksheets/sheet{}.xml", sheet_id_);
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
    }
}

} // namespace core
} // namespace fastexcel
