#include "WorksheetXMLGenerator.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include "fastexcel/core/SharedFormula.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/XMLUtils.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <fmt/format.h>
#include <cstring>
#include <algorithm>
#include <cctype>

namespace fastexcel {
namespace xml {

// 构造函数

WorksheetXMLGenerator::WorksheetXMLGenerator(const core::Worksheet* worksheet)
    : worksheet_(worksheet)
    , workbook_(nullptr)
    , sst_(nullptr)
    , format_repo_(nullptr)
    , mode_(GenerationMode::BATCH) {
    
    if (worksheet_) {
        workbook_ = worksheet_->getParentWorkbook().get();
        if (workbook_) {
            sst_ = workbook_->getSharedStrings();
            format_repo_ = &workbook_->getStyles();
            
            // 根据工作簿模式自动设置生成模式
            auto workbook_mode = workbook_->getOptions().mode;
            if (workbook_mode == core::WorkbookMode::STREAMING) {
                mode_ = GenerationMode::STREAMING;
            }
        }
    }
}

// 主要生成方法

void WorksheetXMLGenerator::generate(const std::function<void(const char*, size_t)>& callback) {
    if (!worksheet_) {
        FASTEXCEL_LOG_ERROR("WorksheetXMLGenerator::generate - worksheet is null");
        return;
    }
    
    if (mode_ == GenerationMode::STREAMING) {
        generateStreaming(callback);
    } else {
        generateBatch(callback);
    }
}

void WorksheetXMLGenerator::generateRelationships(const std::function<void(const char*, size_t)>& callback) {
    if (!worksheet_) return;
    
    // 使用专用的Relationships类生成XML（与Worksheet原来逻辑保持一致）
    xml::Relationships relationships;
    
    // 添加超链接关系
    auto [max_row, max_col] = worksheet_->getUsedRange();
    for (int row = 0; row <= max_row; ++row) {
        for (int col = 0; col <= max_col; ++col) {
            if (worksheet_->hasCellAt(row, col)) {
                const auto& cell = worksheet_->getCell(row, col);
                if (cell.hasHyperlink()) {
                    relationships.addAutoRelationship(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                        cell.getHyperlink(),
                        "External"
                    );
                }
            }
        }
    }
    
    // 添加图片关系（如果有图片）
    const auto& images = worksheet_->getImages();
    if (!images.empty()) {
        std::string drawing_target = "../drawings/drawing" + std::to_string(worksheet_->getSheetId()) + ".xml";
        relationships.addAutoRelationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            drawing_target
        );
    }
    
    // 如果没有关系，不生成任何内容
    if (relationships.size() == 0) {
        return;
    }
    
    // 使用Relationships类生成XML到回调
    relationships.generate(callback);
}

// 批量模式生成方法

void WorksheetXMLGenerator::generateBatch(const std::function<void(const char*, size_t)>& callback) {
    // 使用XMLStreamWriter来正确生成XML
    XMLStreamWriter writer(callback);
    
    writer.startDocument();
    writer.startElement("worksheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 尺寸信息（考虑合并区域扩展列/行）
    auto [max_row, max_col] = worksheet_->getUsedRange();
    // 扩展到合并区域的最远行列
    int dim_last_row = max_row;
    int dim_last_col = max_col;
    const auto& merge_ranges_for_dim = worksheet_->getMergeRanges();
    for (const auto& rng : merge_ranges_for_dim) {
        dim_last_row = std::max(dim_last_row, rng.last_row);
        dim_last_col = std::max(dim_last_col, rng.last_col);
    }
    writer.startElement("dimension");
    if (dim_last_row >= 0 && dim_last_col >= 0) {
        std::string ref = utils::CommonUtils::cellReference(0, 0) + ":" +
                         utils::CommonUtils::cellReference(dim_last_row, dim_last_col);
        writer.writeAttribute("ref", ref.c_str());
    } else {
        writer.writeAttribute("ref", "A1");
    }
    writer.endElement(); // dimension
    
    // 生成工作表视图
    generateSheetViews(writer);
    
    // 工作表格式信息：补充 baseColWidth 与 defaultColWidth 以贴近Excel行为
    writer.startElement("sheetFormatPr");
    // 基础列宽字符数（通常Excel为10，对应Calibri 11的最大数字字符宽度计数）
    writer.writeAttribute("baseColWidth", "10");
    writer.writeAttribute("defaultRowHeight", "15");
    // 使用工作表的默认列宽（Worksheet为friend，可直接访问）
    writer.writeAttribute("defaultColWidth", fmt::format("{}", worksheet_->default_col_width_).c_str());
    writer.endElement(); // sheetFormatPr
    
    // 生成列信息
    generateColumns(writer);
    
    // 生成单元格数据
    generateSheetData(writer);
    
    // 生成合并单元格
    generateMergeCells(writer);

    // 注意元素顺序：根据 SpreadsheetML schema，sheetProtection 应位于 autoFilter 之前。
    // 先输出工作表保护，再输出自动筛选，避免 Excel 修复提示。
    generateSheetProtection(writer);

    // 生成自动筛选
    generateAutoFilter(writer);
    
    // 生成图片绘图引用
    generateDrawing(writer);
    
    writer.endElement(); // worksheet
    writer.endDocument();
}

void WorksheetXMLGenerator::generateSheetViews(XMLStreamWriter& writer) {
    writer.startElement("sheetViews");
    writer.startElement("sheetView");
    
    if (worksheet_->isTabSelected()) {
        writer.writeAttribute("tabSelected", "1");
    }
    writer.writeAttribute("workbookViewId", "0");
    
    // 缩放比例
    if (worksheet_->getZoom() != 100) {
        writer.writeAttribute("zoomScale", fmt::format("{}", worksheet_->getZoom()).c_str());
    }
    
    // 网格线
    if (!worksheet_->isGridlinesVisible()) {
        writer.writeAttribute("showGridLines", "0");
    }
    
    // 行列标题
    if (!worksheet_->isRowColHeadersVisible()) {
        writer.writeAttribute("showRowColHeaders", "0");
    }
    
    // 从右到左
    if (worksheet_->isRightToLeft()) {
        writer.writeAttribute("rightToLeft", "1");
    }
    
    // 冻结窗格
    if (worksheet_->hasFrozenPanes()) {
        auto freeze_info = worksheet_->getFreezeInfo();
        writer.startElement("pane");
        if (freeze_info.col > 0) {
            writer.writeAttribute("xSplit", fmt::format("{}", freeze_info.col).c_str());
        }
        if (freeze_info.row > 0) {
            writer.writeAttribute("ySplit", fmt::format("{}", freeze_info.row).c_str());
        }
        
        // Excel 常见行为：冻结首行时 topLeftCell = A2；冻结首列时 = B1；同时冻结 = B2
        int top_left_row = freeze_info.row > 0 ? freeze_info.row : 0;
        int top_left_col = freeze_info.col > 0 ? freeze_info.col : 0;
        std::string top_left = utils::CommonUtils::cellReference(top_left_row, top_left_col);
        writer.writeAttribute("topLeftCell", top_left.c_str());
        
        writer.writeAttribute("state", "frozen");
        writer.endElement(); // pane
    }
    
    writer.endElement(); // sheetView
    writer.endElement(); // sheetViews
}

void WorksheetXMLGenerator::generateColumns(XMLStreamWriter& writer) {
    const auto& col_info = worksheet_->getColumnInfo();
    if (col_info.empty()) return;
    
    writer.startElement("cols");
    
    // 排序列信息以便合并相邻的相同属性列
    std::vector<std::pair<int, core::ColumnInfo>> sorted_columns(col_info.begin(), col_info.end());
    std::sort(sorted_columns.begin(), sorted_columns.end());
    
    for (size_t i = 0; i < sorted_columns.size(); ) {
        int min_col = sorted_columns[i].first;
        const auto& info = sorted_columns[i].second;
        int max_col = min_col;
        
        // 查找具有相同属性的连续列
        while (i + 1 < sorted_columns.size() && 
               sorted_columns[i + 1].first == max_col + 1 &&
               sorted_columns[i + 1].second == info) {
            max_col = sorted_columns[i + 1].first;
            ++i;
        }
        
        // 生成合并后的<col>标签
        writer.startElement("col");
        writer.writeAttribute("min", fmt::format("{}", min_col + 1).c_str());
        writer.writeAttribute("max", fmt::format("{}", max_col + 1).c_str());
        
        if (info.width > 0) {
            writer.writeAttribute("width", fmt::format("{}", info.width).c_str());
            writer.writeAttribute("customWidth", "1");
        }
        
        if (info.format_id >= 0) {
            writer.writeAttribute("style", fmt::format("{}", info.format_id).c_str());
        }
        
        if (info.hidden) {
            writer.writeAttribute("hidden", "1");
        }
        
        writer.endElement(); // col
        ++i;
    }
    
    writer.endElement(); // cols
}

// 其他方法

void WorksheetXMLGenerator::generateSheetData(XMLStreamWriter& writer) {
    writer.startElement("sheetData");
    
    auto [max_row, max_col] = worksheet_->getUsedRange();
    if (max_row < 0 || max_col < 0) {
        writer.endElement(); // sheetData (empty)
        return;
    }
    
    // 逐行生成数据
    for (int row = 0; row <= max_row; ++row) {
        bool row_has_data = false;
        
        // 检查行是否有数据
        for (int col = 0; col <= max_col; ++col) {
            if (worksheet_->hasCellAt(row, col)) {
                row_has_data = true;
                break;
            }
        }
        
        if (!row_has_data) continue;
        
        writer.startElement("row");
        writer.writeAttribute("r", fmt::format("{}", row + 1).c_str());
        
        for (int col = 0; col <= max_col; ++col) {
            if (!worksheet_->hasCellAt(row, col)) continue;
            
            const auto& cell = worksheet_->getCell(row, col);
            if (cell.isEmpty() && !cell.hasFormat()) continue;
            
            writer.startElement("c");
            writer.writeAttribute("r", utils::CommonUtils::cellReference(row, col).c_str());
            
            // 应用单元格格式
            int format_index = getCellFormatIndex(cell);
            if (format_index >= 0) {
                writer.writeAttribute("s", fmt::format("{}", format_index).c_str());
            }
            
            // 输出单元格值
            if (!cell.isEmpty()) {
                if (cell.isFormula()) {
                    // 检查是否为共享公式
                    if (cell.isSharedFormula()) {
                        int shared_index = cell.getSharedFormulaIndex();
                        auto* shared_formula_manager = worksheet_->getSharedFormulaManager();
                        
                        if (shared_formula_manager) {
                            const core::SharedFormula* shared_formula = shared_formula_manager->getSharedFormula(shared_index);
                            
                            if (shared_formula) {
                                // 检查是否为共享公式的主单元格（第一个单元格）
                                int range_first_row = shared_formula->getRefFirstRow();
                                int range_first_col = shared_formula->getRefFirstCol();
                                int range_last_row = shared_formula->getRefLastRow();
                                int range_last_col = shared_formula->getRefLastCol();
                                bool is_master_cell = (row == range_first_row && col == range_first_col);
                                
                                writer.startElement("f");
                                writer.writeAttribute("t", "shared");
                                writer.writeAttribute("si", fmt::format("{}", shared_index).c_str());
                                
                                if (is_master_cell) {
                                    // 主单元格：输出完整公式和范围
                                    std::string range_ref = utils::CommonUtils::cellReference(range_first_row, range_first_col) + ":" +
                                                          utils::CommonUtils::cellReference(range_last_row, range_last_col);
                                    writer.writeAttribute("ref", range_ref.c_str());
                                    writer.writeText(cell.getFormula());
                                }
                                // 其他单元格：只输出 si 属性，无内容
                                
                                writer.endElement(); // f
                                
                                // ⚠️ 修复：不输出共享公式缓存值，让Excel重新计算
                                // Excel打开文件时会自动计算所有公式，避免缓存值不一致导致"需要修复"提示
                                // double result = cell.getFormulaResult();
                                // writer.startElement("v");
                                // writer.writeText(fmt::format("{}", result).c_str());
                                // writer.endElement(); // v
                            }
                        }
                    } else {
                        // 普通公式 - 不应该设置 t="str" 属性
                        writer.startElement("f");
                        writer.writeText(cell.getFormula());
                        writer.endElement(); // f
                        
                        // ⚠️ 修复：不输出普通公式缓存值，让Excel重新计算
                        // Excel打开文件时会自动计算所有公式，避免缓存值不一致导致"需要修复"提示
                        // double result = cell.getFormulaResult();
                        // writer.startElement("v");
                        // writer.writeText(fmt::format("{}", result).c_str());
                        // writer.endElement(); // v
                    }
                } else if (cell.isString()) {
                    if (workbook_ && workbook_->getOptions().use_shared_strings && sst_) {
                        writer.writeAttribute("t", "s");
                        writer.startElement("v");
                        int sst_index = sst_->getStringId(cell.getStringValue());
                        if (sst_index < 0) {
                            sst_index = const_cast<core::SharedStringTable*>(sst_)->addString(cell.getStringValue());
                        }
                        writer.writeText(fmt::format("{}", sst_index).c_str());
                        writer.endElement(); // v
                    } else {
                        writer.writeAttribute("t", "inlineStr");
                        writer.startElement("is");
                        writer.startElement("t");
                        writer.writeText(utils::XMLUtils::escapeXML(cell.getValue<std::string>()).c_str());
                        writer.endElement(); // t
                        writer.endElement(); // is
                    }
                } else if (cell.isNumber()) {
                    writer.startElement("v");
                    writer.writeText(fmt::format("{}", cell.getNumberValue()).c_str());
                    writer.endElement(); // v
                } else if (cell.isBoolean()) {
                    writer.writeAttribute("t", "b");
                    writer.startElement("v");
                    writer.writeText(cell.getBooleanValue() ? "1" : "0");
                    writer.endElement(); // v
                }
            }
            
            writer.endElement(); // c
        }
        
        writer.endElement(); // row
    }
    
    writer.endElement(); // sheetData
}

void WorksheetXMLGenerator::generateMergeCells(XMLStreamWriter& writer) {
    const auto& merge_ranges = worksheet_->getMergeRanges();
    if (merge_ranges.empty()) return;
    
    writer.startElement("mergeCells");
    writer.writeAttribute("count", fmt::format("{}", merge_ranges.size()).c_str());
    
    for (const auto& range : merge_ranges) {
        writer.startElement("mergeCell");
        std::string range_ref = utils::CommonUtils::cellReference(range.first_row, range.first_col) + ":" +
                               utils::CommonUtils::cellReference(range.last_row, range.last_col);
        writer.writeAttribute("ref", range_ref.c_str());
        writer.endElement(); // mergeCell
    }
    
    writer.endElement(); // mergeCells
}

void WorksheetXMLGenerator::generateAutoFilter(XMLStreamWriter& writer) {
    if (!worksheet_->hasAutoFilter()) return;
    
    auto filter_range = worksheet_->getAutoFilterRange();
    writer.startElement("autoFilter");
    std::string ref = utils::CommonUtils::cellReference(filter_range.first_row, filter_range.first_col) + ":" +
                     utils::CommonUtils::cellReference(filter_range.last_row, filter_range.last_col);
    writer.writeAttribute("ref", ref.c_str());
    writer.endElement(); // autoFilter
}

namespace {
// Excel 工作表保护密码哈希（与 Excel/Libxlsxwriter 一致的16位算法）
static unsigned short hashExcelPassword(const std::string& pwd) {
    if (pwd.empty()) return 0;
    unsigned short hash = 0;
    // 从末尾到开头处理字符，右移并轮转最低位，再异或字符值
    for (int i = static_cast<int>(pwd.size()) - 1; i >= 0; --i) {
        unsigned short low_bit = (hash & 0x0001);
        hash >>= 1;
        if (low_bit) hash |= 0x8000;
        hash ^= static_cast<unsigned char>(pwd[static_cast<size_t>(i)]);
    }
    hash ^= 0xCE4B;
    return hash;
}

static bool isFourHex(const std::string& s) {
    if (s.size() != 4) return false;
    for (char c : s) {
        if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))) return false;
    }
    return true;
}
}

void WorksheetXMLGenerator::generateSheetProtection(XMLStreamWriter& writer) {
    if (!worksheet_->isProtected()) return;

    writer.startElement("sheetProtection");
    writer.writeAttribute("sheet", "1");

    const std::string& password = worksheet_->getProtectionPassword();
    if (!password.empty()) {
        // 如果传入的是4位十六进制，认为已是哈希；否则计算哈希。
        if (isFourHex(password)) {
            // 规范化为大写
            std::string upper = password;
            for (auto& ch : upper) ch = static_cast<char>(std::toupper(static_cast<unsigned char>(ch)));
            writer.writeAttribute("password", upper.c_str());
        } else {
            unsigned short hash = hashExcelPassword(password);
            writer.writeAttribute("password", fmt::format("{:04X}", hash).c_str());
        }
    }

    writer.endElement(); // sheetProtection
}

void WorksheetXMLGenerator::generateDrawing(XMLStreamWriter& writer) {
    // 检查工作表是否有图片
    if (!worksheet_->hasImages()) {
        return;
    }
    
    // 生成drawing引用
    writer.startElement("drawing");
    // drawing 关系通常在 worksheet 关系文件中位于超链接之后
    int hyperlink_count = 0;
    auto [max_row, max_col] = worksheet_->getUsedRange();
    for (int row = 0; row <= max_row; ++row) {
        for (int col = 0; col <= max_col; ++col) {
            if (worksheet_->hasCellAt(row, col)) {
                const auto& cell = worksheet_->getCell(row, col);
                if (cell.hasHyperlink()) {
                    hyperlink_count++;
                }
            }
        }
    }
    
    // drawing 的 rId 应为超链接数量 + 1
    std::string drawing_rel = fmt::format("rId{}", hyperlink_count + 1);
    writer.writeAttribute("r:id", drawing_rel.c_str());
    writer.endElement(); // drawing
}

// 流式模式生成方法

void WorksheetXMLGenerator::generateStreaming(const std::function<void(const char*, size_t)>& callback) {
    // 使用XMLStreamWriter替代字符串拼接，提高性能
    XMLStreamWriter writer(callback);
    
    writer.startDocument();
    writer.startElement("worksheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 尺寸信息（考虑合并区域扩展列/行）
    auto [max_row2, max_col2] = worksheet_->getUsedRange();
    int dim_last_row2 = max_row2;
    int dim_last_col2 = max_col2;
    const auto& merge_ranges_stream = worksheet_->getMergeRanges();
    for (const auto& rng : merge_ranges_stream) {
        dim_last_row2 = std::max(dim_last_row2, rng.last_row);
        dim_last_col2 = std::max(dim_last_col2, rng.last_col);
    }
    writer.startElement("dimension");
    if (dim_last_row2 >= 0 && dim_last_col2 >= 0) {
        std::string ref2 = utils::CommonUtils::cellReference(0, 0) + ":" +
                          utils::CommonUtils::cellReference(dim_last_row2, dim_last_col2);
        writer.writeAttribute("ref", ref2.c_str());
    } else {
        writer.writeAttribute("ref", "A1");
    }
    writer.endElement(); // dimension
    
    // 简化的工作表视图
    writer.startElement("sheetViews");
    writer.startElement("sheetView");
    writer.writeAttribute("workbookViewId", "0");
    if (worksheet_->isTabSelected()) {
        writer.writeAttribute("tabSelected", "1");
    }
    writer.endElement(); // sheetView
    writer.endElement(); // sheetViews
    
    // 工作表格式信息（流式）：同样补充 baseColWidth 与 defaultColWidth
    writer.startElement("sheetFormatPr");
    writer.writeAttribute("baseColWidth", "10");
    writer.writeAttribute("defaultRowHeight", "15");
    writer.writeAttribute("defaultColWidth", fmt::format("{}", worksheet_->default_col_width_).c_str());
    writer.endElement(); // sheetFormatPr
    
    // 列信息（使用XMLStreamWriter优化）
    const auto& col_info = worksheet_->getColumnInfo();
    if (!col_info.empty()) {
        writer.startElement("cols");
        for (const auto& [col_num, info] : col_info) {
            writer.startElement("col");
            writer.writeAttribute("min", fmt::format("{}", col_num + 1).c_str());
            writer.writeAttribute("max", fmt::format("{}", col_num + 1).c_str());
            if (info.width > 0) {
                writer.writeAttribute("width", fmt::format("{}", info.width).c_str());
                writer.writeAttribute("customWidth", "1");
            }
            if (info.hidden) {
                writer.writeAttribute("hidden", "1");
            }
            writer.endElement(); // col
        }
        writer.endElement(); // cols
    }
    
    // 单元格数据（流式处理）
    writer.startElement("sheetData");
    if (max_row2 >= 0 && max_col2 >= 0) {
        generateSheetDataStreaming(writer);  // 传递writer引用
    }
    writer.endElement(); // sheetData
    
    // 合并单元格（使用XMLStreamWriter优化）
    const auto& merge_ranges = worksheet_->getMergeRanges();
    if (!merge_ranges.empty()) {
        writer.startElement("mergeCells");
        writer.writeAttribute("count", fmt::format("{}", merge_ranges.size()).c_str());
        
        for (const auto& range : merge_ranges) {
            writer.startElement("mergeCell");
            std::string range_ref = utils::CommonUtils::cellReference(range.first_row, range.first_col) + ":" +
                                   utils::CommonUtils::cellReference(range.last_row, range.last_col);
            writer.writeAttribute("ref", range_ref.c_str());
            writer.endElement(); // mergeCell
        }
        
        writer.endElement(); // mergeCells
    }
    // 注意元素顺序一致性：先 sheetProtection 再 autoFilter
    generateSheetProtection(writer);
    generateAutoFilter(writer);

    // 生成图片绘图引用（流式模式）
    generateDrawing(writer);
    
    writer.endElement(); // worksheet
    writer.endDocument();
}

void WorksheetXMLGenerator::generateSheetDataStreaming(XMLStreamWriter& writer) {
    auto [max_row, max_col] = worksheet_->getUsedRange();
    if (max_row < 0 || max_col < 0) return;
    
    // 不需要创建新的XMLStreamWriter，使用传入的writer
    
    // 分块处理：每次处理一定数量的行
    const int CHUNK_SIZE = 1000;
    
    for (int chunk_start = 0; chunk_start <= max_row; chunk_start += CHUNK_SIZE) {
        int chunk_end = std::min(chunk_start + CHUNK_SIZE - 1, max_row);
        
        for (int row_num = chunk_start; row_num <= chunk_end; ++row_num) {
            // 检查这一行是否有数据
            bool has_data = false;
            for (int col = 0; col <= max_col; ++col) {
                if (worksheet_->hasCellAt(row_num, col)) {
                    has_data = true;
                    break;
                }
            }
            
            if (!has_data) continue;
            
            // 生成行开始标签
            writer.startElement("row");
            writer.writeAttribute("r", fmt::format("{}", row_num + 1).c_str());
            
            // 生成单元格数据
            for (int col = 0; col <= max_col; ++col) {
                if (!worksheet_->hasCellAt(row_num, col)) continue;
                
                const auto& cell = worksheet_->getCell(row_num, col);
                if (cell.isEmpty() && !cell.hasFormat()) continue;
                
                // 使用优化后的单元格生成方法
                generateCellXMLStreaming(writer, row_num, col, cell);
            }
            
            // 行结束标签
            writer.endElement(); // row
        }
    }
}

// 辅助方法

int WorksheetXMLGenerator::getCellFormatIndex(const core::Cell& cell) {
    if (!cell.hasFormat() || !format_repo_) {
        return -1;
    }
    
    auto format_descriptor = cell.getFormatDescriptor();
    if (!format_descriptor) {
        return -1;
    }
    
    // 查找格式在仓库中的索引
    for (size_t i = 0; i < format_repo_->getFormatCount(); ++i) {
        auto stored_format = format_repo_->getFormat(static_cast<int>(i));
        if (stored_format && *stored_format == *format_descriptor) {
            return static_cast<int>(i);
        }
    }
    
    return -1;
}

void WorksheetXMLGenerator::generateCellXMLStreaming(XMLStreamWriter& writer, int row, int col, const core::Cell& cell) {
    writer.startElement("c");
    writer.writeAttribute("r", utils::CommonUtils::cellReference(row, col).c_str());
    
    // 应用单元格格式
    int format_index = getCellFormatIndex(cell);
    if (format_index >= 0) {
        writer.writeAttribute("s", fmt::format("{}", format_index).c_str());
    }
    
    // 输出单元格值
    if (!cell.isEmpty()) {
        if (cell.isFormula()) {
            // 检查是否为共享公式
            if (cell.isSharedFormula()) {
                int shared_index = cell.getSharedFormulaIndex();
                auto* shared_formula_manager = worksheet_->getSharedFormulaManager();
                
                if (shared_formula_manager) {
                    const core::SharedFormula* shared_formula = shared_formula_manager->getSharedFormula(shared_index);
                    
                    if (shared_formula) {
                        // 检查是否为共享公式的主单元格
                        int range_first_row = shared_formula->getRefFirstRow();
                        int range_first_col = shared_formula->getRefFirstCol();
                        int range_last_row = shared_formula->getRefLastRow();
                        int range_last_col = shared_formula->getRefLastCol();
                        bool is_master_cell = (row == range_first_row && col == range_first_col);
                        
                        writer.startElement("f");
                        writer.writeAttribute("t", "shared");
                        writer.writeAttribute("si", fmt::format("{}", shared_index).c_str());
                        
                        if (is_master_cell) {
                            // 主单元格：输出完整公式和范围
                            std::string range_ref = utils::CommonUtils::cellReference(range_first_row, range_first_col) + ":" +
                                                  utils::CommonUtils::cellReference(range_last_row, range_last_col);
                            writer.writeAttribute("ref", range_ref.c_str());
                            writer.writeText(cell.getFormula());
                        }
                        // 其他单元格：只输出 si 属性，无内容
                        
                        writer.endElement(); // f
                        
                        // 输出结果值 (对于共享公式，总是输出结果，即使是0)
                        double result = cell.getFormulaResult();
                        writer.startElement("v");
                        writer.writeText(fmt::format("{}", result).c_str());
                        writer.endElement(); // v
                    }
                }
            } else {
                // 普通公式 - 不应该设置 t="str" 属性 
                writer.startElement("f");
                writer.writeText(cell.getFormula());
                writer.endElement(); // f
                
                // ⚠️ 修复：不输出公式缓存值，让Excel重新计算
                // Excel打开文件时会自动计算所有公式，避免缓存值不一致导致"需要修复"提示
                // double result = cell.getFormulaResult();
                // writer.startElement("v");
                // writer.writeText(fmt::format("{}", result).c_str());
                // writer.endElement(); // v
            }
        } else if (cell.isString()) {
            if (workbook_ && workbook_->getOptions().use_shared_strings && sst_) {
                writer.writeAttribute("t", "s");
                writer.startElement("v");
                int sst_index = sst_->getStringId(cell.getStringValue());
                if (sst_index < 0) {
                    sst_index = const_cast<core::SharedStringTable*>(sst_)->addString(cell.getStringValue());
                }
                writer.writeText(fmt::format("{}", sst_index).c_str());
                writer.endElement(); // v
            } else {
                writer.writeAttribute("t", "inlineStr");
                writer.startElement("is");
                writer.startElement("t");
                writer.writeText(utils::XMLUtils::escapeXML(cell.getValue<std::string>()).c_str());
                writer.endElement(); // t
                writer.endElement(); // is
            }
        } else if (cell.isNumber()) {
            writer.startElement("v");
            writer.writeText(fmt::format("{}", cell.getNumberValue()).c_str());
            writer.endElement(); // v
        } else if (cell.isBoolean()) {
            writer.writeAttribute("t", "b");
            writer.startElement("v");
            writer.writeText(cell.getBooleanValue() ? "1" : "0");
            writer.endElement(); // v
        }
    }
    
    writer.endElement(); // c
}

}} // namespace fastexcel::xml
