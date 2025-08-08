#include "fastexcel/opc/PackageEditor.hpp"
#include "fastexcel/opc/ZipRepackWriter.hpp"
#include "fastexcel/opc/PartGraph.hpp"
#include "fastexcel/archive/ZipReader.hpp"  // Add missing include
#include "fastexcel/archive/ZipArchive.hpp" // For ZipError enum
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/xml/StyleSerializer.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include <sstream>
#include <algorithm>
#include <cmath>
#include <cctype>       // 添加用于isdigit函数

namespace fastexcel {
namespace opc {

// ========== CellRef 实现 ==========

std::string PackageEditor::CellRef::toString() const {
    // 使用CommonUtils的方法生成单元格引用
    return utils::CommonUtils::cellReference(row, col);
}

// ========== 简化的数据模型定义 ==========

struct PackageEditor::WorkbookModel {
    std::vector<std::string> sheet_names;
    // 其他workbook相关数据
};

struct PackageEditor::WorksheetModel {
    std::string name;
    // 单元格数据等
};

struct PackageEditor::StylesModel {
    // 样式数据
};

struct PackageEditor::SharedStringsModel {
    std::vector<std::string> strings;
    std::unordered_map<std::string, size_t> string_index;
};

// ========== EditPlan实现 ==========

void PackageEditor::EditPlan::markDirty(const std::string& part) {
    dirty_parts.insert(part);
    markRelatedDirty(part);
}

void PackageEditor::EditPlan::markRelatedDirty(const std::string& part) {
    // 智能标记关联部件
    
    // 工作表修改 -> 可能影响calcChain
    if (part.find("xl/worksheets/sheet") == 0) {
        removed_parts.insert("xl/calcChain.xml");  // 保守策略：删除calcChain
    }
    
    // 共享字符串修改 -> workbook.xml.rels可能需要更新
    if (part == "xl/sharedStrings.xml") {
        dirty_parts.insert("xl/_rels/workbook.xml.rels");
    }
    
    // 新增工作表 -> 影响workbook.xml和Content_Types
    if (part.find("xl/worksheets/sheet") == 0 && new_parts.count(part)) {
        dirty_parts.insert("xl/workbook.xml");
        dirty_parts.insert("[Content_Types].xml");
        dirty_parts.insert("xl/_rels/workbook.xml.rels");
    }
    
    // 样式修改 -> 可能影响所有工作表（但第一期不处理）
    if (part == "xl/styles.xml") {
        // 保守：可以标记所有sheet为需要验证
    }
}

void PackageEditor::EditPlan::addNewPart(const std::string& path, const std::string& content) {
    new_parts[path] = content;
    markDirty(path);
}

void PackageEditor::EditPlan::removePart(const std::string& path) {
    removed_parts.insert(path);
    new_parts.erase(path);
    dirty_parts.erase(path);
}

// ========== 工厂方法 ==========

std::unique_ptr<PackageEditor> PackageEditor::open(const core::Path& xlsx_path) {
    auto editor = std::unique_ptr<PackageEditor>(new PackageEditor());
    
    if (!editor->initialize(xlsx_path)) {
        LOG_ERROR("Failed to initialize PackageEditor from: {}", xlsx_path.string());
        return nullptr;
    }
    
    LOG_INFO("Opened Excel package for editing: {}", xlsx_path.string());
    return editor;
}

std::unique_ptr<PackageEditor> PackageEditor::fromWorkbook(core::Workbook* workbook) {
    if (!workbook) {
        LOG_ERROR("Cannot create PackageEditor from null Workbook");
        return nullptr;
    }
    
    auto editor = std::unique_ptr<PackageEditor>(new PackageEditor());
    editor->workbook_ = workbook;
    editor->owns_workbook_ = false;  // 不拥有外部传入的workbook
    
    if (!editor->initializeFromWorkbook()) {
        LOG_ERROR("Failed to initialize PackageEditor from Workbook");
        return nullptr;
    }
    
    LOG_INFO("Created PackageEditor from existing Workbook with {} sheets", 
             workbook->getWorksheetNames().size());
    return editor;
}

std::unique_ptr<PackageEditor> PackageEditor::create() {
    auto editor = std::unique_ptr<PackageEditor>(new PackageEditor());
    
    // 先创建一个临时的 Workbook
    core::Path temp_path("temp_new_workbook.xlsx");
    auto temp_workbook = std::make_unique<core::Workbook>(temp_path);
    if (!temp_workbook->open()) {
        LOG_ERROR("无法创建临时 Workbook");
        return nullptr;
    }
    
    // 添加默认工作表
    temp_workbook->addWorksheet("Sheet1");
    
    // 设置 Workbook 对象
    editor->workbook_ = temp_workbook.release();  // 转移所有权
    editor->owns_workbook_ = true;  // 我们拥有这个workbook的所有权
    
    if (!editor->initializeFromWorkbook()) {
        LOG_ERROR("Failed to initialize empty PackageEditor");
        return nullptr;
    }
    
    LOG_INFO("Created new Excel package with default sheet");
    return editor;
}

// ========== 构造/析构 ==========

PackageEditor::PackageEditor() {
}

PackageEditor::~PackageEditor() {
    if (isDirty()) {
        LOG_WARN("PackageEditor destroyed with {} unsaved changes", edit_plan_.dirty_parts.size());
    }
    
    // 清理 workbook_ 对象（如果我们拥有它的话）
    // 注意：在 fromWorkbook 情况下，我们不拥有 workbook_ 的所有权
    // 只有在 create() 情况下我们才拥有所有权
    if (workbook_ && owns_workbook_) {
        delete workbook_;
        workbook_ = nullptr;
    }
}

// ========== 初始化 ==========

bool PackageEditor::initialize(const core::Path& xlsx_path) {
    source_path_ = xlsx_path;
    
    // 1. 打开源ZIP（只读）
    source_reader_ = std::make_unique<archive::ZipReader>(xlsx_path);
    if (!source_reader_->open()) {
        LOG_ERROR("Failed to open source ZIP: {}", xlsx_path.string());
        return false;
    }
    
    // 2. 构建部件关系图
    part_graph_ = std::make_unique<PartGraph>();
    // Use archive::ZipReader interface
    if (!part_graph_->buildFromZipReader(source_reader_.get())) {
        LOG_ERROR("Failed to build part graph");
        return false;
    }
    
    // 3. 解析Content_Types
    content_types_ = std::make_unique<ContentTypes>();
    std::string content_types_xml;
    // Use extractFile instead of readEntry
    if (source_reader_->extractFile("[Content_Types].xml", content_types_xml) == archive::ZipError::Ok) {
        content_types_->parse(content_types_xml);
    }
    
    LOG_DEBUG("Initialized with {} parts", part_graph_->getAllParts().size());
    return true;
}

bool PackageEditor::initializeFromWorkbook() {
    // 从 workbook_ 初始化基本结构
    part_graph_ = std::make_unique<PartGraph>();
    content_types_ = std::make_unique<ContentTypes>();
    
    // 添加基本的 Content Types
    content_types_->addDefault("rels", "application/vnd.openxmlformats-package.relationships+xml");
    content_types_->addDefault("xml", "application/xml");
    
    // 添加核心覆盖类型
    content_types_->addOverride("/xl/workbook.xml", 
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    content_types_->addOverride("/xl/styles.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
    
    // 根据 Workbook 内容标记相应部件为 dirty
    edit_plan_.markDirty("xl/workbook.xml");
    edit_plan_.markDirty("xl/styles.xml");
    edit_plan_.markDirty("[Content_Types].xml");
    edit_plan_.markDirty("_rels/.rels");
    edit_plan_.markDirty("xl/_rels/workbook.xml.rels");
    
    // 为每个工作表添加内容类型和标记为dirty
    auto sheet_names = workbook_->getWorksheetNames();
    for (size_t i = 0; i < sheet_names.size(); ++i) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(i + 1) + ".xml";
        edit_plan_.markDirty(sheet_path);
        
        // 添加工作表的内容类型
        content_types_->addOverride("/" + sheet_path,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    }
    
    // 如果使用共享字符串，标记为dirty
    if (workbook_->getOptions().use_shared_strings) {
        edit_plan_.markDirty("xl/sharedStrings.xml");
        content_types_->addOverride("/xl/sharedStrings.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
    }
    
    LOG_DEBUG("Initialized from Workbook with {} sheets, {} dirty parts", 
              sheet_names.size(), edit_plan_.dirty_parts.size());
    return true;
}

    // 只留下必要的初始化方法

// ========== 核心编辑API ==========

void PackageEditor::setCell(const SheetId& sheet, const CellRef& ref, const CellValue& value) {
    // 验证单元格位置
    if (!isValidCellRef(ref.row, ref.col)) {
        LOG_ERROR("无效的单元格引用：行 {} 列 {}", ref.row, ref.col);
        return;
    }
    
    // 验证工作表名称
    if (sheet.empty()) {
        LOG_ERROR("工作表名称不能为空");
        return;
    }
    
    // 确保工作表存在
    if (getSheetId(sheet) <= 0) {
        LOG_ERROR("工作表 '{}' 不存在", sheet);
        return;
    }
    
    // 确保工作表模型已加载
    ensureSheetModel(sheet);
    
    // 更新模型
    // sheet_models_[sheet]->setCell(ref, value);
    
    // 获取工作表路径
    std::string sheet_path = getSheetPath(sheet);
    if (sheet_path.empty()) {
        LOG_ERROR("无法获取工作表路径：{}", sheet);
        return;
    }
    
    // 标记工作表为dirty
    edit_plan_.markDirty(sheet_path);
    
    // 如果使用共享字符串
    if (options_.use_shared_strings && value.type == CellValue::String) {
        ensureSharedStringsModel();
        // sst_model_->addString(value.str_value);
        edit_plan_.markDirty("xl/sharedStrings.xml");
    }
    
    LOG_DEBUG("Set cell {} in sheet {} (path: {})", ref.toString(), sheet, sheet_path);
}

void PackageEditor::addSheet(const std::string& name) {
    // 验证工作表名称
    if (!isValidSheetName(name)) {
        LOG_ERROR("无效的工作表名称：'{}'", name);
        return;
    }
    
    // 必须有工作簿对象才能添加工作表
    if (!workbook_) {
        LOG_ERROR("无法在没有 Workbook 对象的情况下添加工作表");
        return;
    }
    
    // 检查工作表名称是否已存在
    if (getSheetId(name) > 0) {
        LOG_WARN("工作表 '{}' 已存在，跳过添加", name);
        return;
    }
    
    // 检查工作表数量限制
    auto existing_sheets = getSheetNames();
    if (existing_sheets.size() >= 255) {  // Excel 的最大工作表数量
        LOG_ERROR("已达到最大工作表数量限制 (255)");
        return;
    }
    
    // 通过Workbook添加新工作表
    try {
        workbook_->addWorksheet(name);
        auto worksheet = workbook_->getWorksheet(name);
        if (!worksheet) {
            LOG_ERROR("在 Workbook 中添加工作表失败：{}", name);
            return;
        }
    } catch (const std::exception& e) {
        LOG_ERROR("添加工作表时发生异常：{} - {}", name, e.what());
        return;
    }
    
    // 重新构建映射（因为添加了新工作表）
    rebuildSheetMapping();
    
    // 获取新工作表的路径
    std::string sheet_path = getSheetPath(name);
    if (sheet_path.empty()) {
        LOG_ERROR("无法获取工作表路径：{}", name);
        return;
    }
    
    // 标记为dirty
    edit_plan_.markDirty(sheet_path);
    edit_plan_.markDirty("xl/workbook.xml");
    
    // 更新Content_Types
    content_types_->addOverride("/" + sheet_path,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    edit_plan_.markDirty("[Content_Types].xml");
    
    LOG_INFO("添加新工作表成功：{} (路径：{})", name, sheet_path);
}

// ========== 提交（核心Repack逻辑） ==========

bool PackageEditor::commit(const core::Path& dst) {
    LOG_INFO("Starting commit to: {}", dst.string());
    LOG_INFO("Dirty parts: {}, New parts: {}, Removed parts: {}", 
             edit_plan_.dirty_parts.size(), 
             edit_plan_.new_parts.size(),
             edit_plan_.removed_parts.size());
    
    // 创建Repack写入器
    ZipRepackWriter writer(dst);
    
    // 收集所有需要写入的条目
    std::unordered_set<std::string> written;
    
    // ========== 阶段1：写入所有dirty和new部件 ==========
    LOG_DEBUG("Phase 1: Writing dirty and new parts");
    
    for (const auto& part : edit_plan_.dirty_parts) {
        if (edit_plan_.removed_parts.count(part)) {
            continue;  // 跳过已删除的
        }
        
        std::string content;
        if (edit_plan_.new_parts.count(part)) {
            // 新部件
            content = edit_plan_.new_parts[part];
        } else {
            // 修改的部件，生成新内容
            content = generatePart(part);
        }
        
        if (!writer.add(part, content)) {
            LOG_ERROR("Failed to write part: {}", part);
            return false;
        }
        written.insert(part);
        LOG_DEBUG("Wrote dirty part: {} ({} bytes)", part, content.size());
    }
    
    // ========== 阶段2：复制未修改的部件（懒复制） ==========
    LOG_DEBUG("Phase 2: Copying unchanged parts");
    
    if (source_reader_) {
        // Use listFiles instead of listEntries
        auto all_entries = source_reader_->listFiles();
        std::vector<std::string> to_copy;
        
        for (const auto& entry : all_entries) {
            // 跳过已写入、已删除的
            if (written.count(entry) || edit_plan_.removed_parts.count(entry)) {
                continue;
            }
            
            // calcChain特殊处理
            if (entry == "xl/calcChain.xml" && options_.remove_calc_chain) {
                LOG_DEBUG("Removing calcChain.xml as per options");
                continue;
            }
            
            to_copy.push_back(entry);
        }
        
        // 批量复制（高效）
        if (!to_copy.empty()) {
            LOG_DEBUG("Copying {} unchanged entries", to_copy.size());
            if (!writer.copyBatch(source_reader_.get(), to_copy)) {
                LOG_ERROR("Failed to copy unchanged parts");
                return false;
            }
        }
    }
    
    // ========== 阶段3：确保关键文件存在 ==========
    LOG_DEBUG("Phase 3: Ensuring critical files");
    
    // 确保[Content_Types].xml存在
    if (!writer.hasEntry("[Content_Types].xml")) {
        std::string content_types_xml = content_types_->serialize();
        writer.add("[Content_Types].xml", content_types_xml);
        LOG_DEBUG("Generated [Content_Types].xml");
    }
    
    // 确保关系文件存在
    std::vector<std::string> required_rels = {
        "_rels/.rels",
        "xl/_rels/workbook.xml.rels"
    };
    
    for (const auto& rels : required_rels) {
        if (!writer.hasEntry(rels)) {
            std::string rels_xml = generateRels(rels);
            writer.add(rels, rels_xml);
            LOG_DEBUG("Generated {}", rels);
        }
    }
    
    // ========== 阶段4：完成写入 ==========
    if (!writer.finish()) {
        LOG_ERROR("Failed to finalize ZIP package");
        return false;
    }
    
    auto stats = writer.getStats();
    LOG_INFO("Commit completed: {} entries added, {} entries copied, {} bytes total",
             stats.entries_added, stats.entries_copied, stats.total_size);
    
    // 清空编辑计划
    edit_plan_ = EditPlan();
    
    return true;
}

bool PackageEditor::save() {
    // 新架构：使用流式处理，避免临时文件
    if (source_path_.empty()) {
        LOG_ERROR("No source path specified, use commit() with target path");
        return false;
    }
    
    // 对于新架构，可以直接修改原文件（但为了安全起见，仍然使用临时文件）
    // TODO: 在完全实现ZipRepackWriter的原地修改功能后，可以去除临时文件
    core::Path temp_path(source_path_.string() + ".tmp");
    
    LOG_INFO("Saving changes to: {} (via temp file for safety)", source_path_.string());
    
    if (!commit(temp_path)) {
        return false;
    }
    
    // 原子替换原文件
    try {
        // 关闭源读取器以释放文件句柄
        if (source_reader_) {
            source_reader_->close();
        }
        
        if (source_path_.exists()) {
            source_path_.remove();
        }
        
        if (!temp_path.moveTo(source_path_)) {
            LOG_ERROR("Failed to move temp file to original location");
            return false;
        }
        
        // 重新打开源文件以便后续操作
        if (source_reader_) {
            source_reader_->open();
        }
        
        LOG_INFO("Successfully saved to: {}", source_path_.string());
        return true;
    } catch (const std::exception& e) {
        LOG_ERROR("Failed to replace original file: {}", e.what());
        return false;
    }
}

// ========== 部件生成 ==========

std::string PackageEditor::generatePart(const std::string& path) const {
    LOG_DEBUG("Generating part: {}", path);
    
    if (path == "xl/workbook.xml") {
        return generateWorkbook();
    }
    if (path.find("xl/worksheets/sheet") == 0) {
        LOG_DEBUG("进入工作表路径解析：{}", path);
        
        // 提取sheet名称，使用更安全的解析方式
        // 从路径提取sheet ID："xl/worksheets/sheet3.xml" -> "3"
        
        // 使用正则表达式模式匹配会更精确，但为了保持简单，使用字符串操作
        constexpr const char* prefix = "xl/worksheets/sheet";
        constexpr size_t prefix_len = 19; // strlen("xl/worksheets/sheet") = 19, not 18!
        
        LOG_DEBUG("前缀长度：{}，路径长度：{}", prefix_len, path.length());
        LOG_DEBUG("前缀内容：'{}'", prefix);
        
        // 验证前缀长度
        size_t actual_prefix_len = std::string(prefix).length();
        LOG_DEBUG("实际前缀长度：{}", actual_prefix_len);
        
        // 既然外层已经检查过前缀匹配，这里直接处理
        if (path.length() > actual_prefix_len) {
            LOG_DEBUG("前缀匹配成功");
            size_t xml_pos = path.find(".xml");
            LOG_DEBUG("找到.xml位置：{}", xml_pos);
            
            if (xml_pos != std::string::npos && xml_pos > actual_prefix_len) {
                std::string id_str = path.substr(actual_prefix_len, xml_pos - actual_prefix_len);
                LOG_DEBUG("解析工作表路径：{} -> ID字符串: '{}' (长度: {})", path, id_str, id_str.length());
                
                // 检查每个字符
                for (size_t i = 0; i < id_str.length(); ++i) {
                    char c = id_str[i];
                    LOG_DEBUG("  字符[{}]: '{}' (ASCII: {}) isdigit: {}", i, c, static_cast<int>(c), std::isdigit(c) ? "true" : "false");
                }
                
                // 验证ID字符串只包含数字
                if (!id_str.empty() && std::all_of(id_str.begin(), id_str.end(), ::isdigit)) {
                    try {
                        int sheet_id = std::stoi(id_str);
                        std::string sheet_name = getSheetName(sheet_id);
                        if (!sheet_name.empty()) {
                            LOG_DEBUG("找到工作表ID {} 对应的名称：'{}'", sheet_id, sheet_name);
                            return generateWorksheet(sheet_name);
                        }
                    } catch (const std::exception& e) {
                        LOG_ERROR("解析工作表ID失败：{} - {}", path, e.what());
                    }
                } else {
                    LOG_ERROR("无效的工作表ID格式：{} -> '{}' (长度: {})", path, id_str, id_str.length());
                }
            } else {
                LOG_ERROR("未找到.xml后缀或位置不正确：xml_pos={}, actual_prefix_len={}", xml_pos, actual_prefix_len);
            }
        } else {
            LOG_ERROR("路径长度不足：路径长度={}, actual_prefix_len={}", path.length(), actual_prefix_len);
        }
        
        // 如果无法解析，尝试使用第一个工作表
        auto sheet_names = getSheetNames();
        if (!sheet_names.empty()) {
            LOG_WARN("无法解析工作表路径 {}uff0c使用第一个工作表: '{}'", path, sheet_names[0]);
            return generateWorksheet(sheet_names[0]);
        }
        
        LOG_ERROR("无法解析工作表路径且没有可用的工作表：{}", path);
        return "";
    }
    if (path == "xl/styles.xml") {
        return generateStyles();
    }
    if (path == "xl/sharedStrings.xml") {
        return generateSharedStrings();
    }
    if (path == "[Content_Types].xml") {
        return content_types_->serialize();
    }
    if (path.find(".rels") != std::string::npos) {
        return generateRels(path);
    }
    
    // 默认返回空
    LOG_WARN("No generator for part: {}", path);
    return "";
}

std::string PackageEditor::generateWorkbook() const {
    // 必须有工作簿对象才能生成workbook.xml
    if (!workbook_) {
        LOG_ERROR("Cannot generate workbook.xml without a Workbook object");
        return "";
    }
    
    std::ostringstream buffer;
    auto callback = [&buffer](const char* data, size_t size) {
        buffer.write(data, static_cast<std::streamsize>(size));
    };
    
    // 调用Workbook的generateWorkbookXML方法（现在可以访问，因为PackageEditor是友元类）
    workbook_->generateWorkbookXML(callback);
    return buffer.str();
}

std::string PackageEditor::generateWorksheet(const SheetId& sheet) const {
    // 必须有工作簿对象才能生成worksheet XML
    if (!workbook_) {
        LOG_ERROR("Cannot generate worksheet XML without a Workbook object");
        return "";
    }
    
    auto worksheet = workbook_->getWorksheet(sheet);
    if (!worksheet) {
        LOG_ERROR("Worksheet '{}' not found in workbook", sheet);
        return "";
    }
    
    std::ostringstream buffer;
    auto callback = [&buffer](const char* data, size_t size) {
        buffer.write(data, static_cast<std::streamsize>(size));
    };
    
    // 调用Worksheet的generateXML方法
    worksheet->generateXML(callback);
    return buffer.str();
}

// ========== 辅助方法 ==========

std::vector<std::string> PackageEditor::getSheetNames() const {
    if (workbook_) {
        // 从工作簿获取所有工作表名称
        auto names = workbook_->getWorksheetNames();
        LOG_DEBUG("获取工作表名称列表：{} 个工作表", names.size());
        for (const auto& name : names) {
            LOG_DEBUG("  - '{}'", name);
        }
        return names;
    }
    
    // TODO: 从workbook_model_获取
    return {"Sheet1"};  // 临时实现
}

std::vector<std::string> PackageEditor::getDirtyParts() const {
    return std::vector<std::string>(edit_plan_.dirty_parts.begin(), 
                                   edit_plan_.dirty_parts.end());
}

void PackageEditor::ensureWorkbookModel() const {
    if (workbook_model_) return;
    
    // 懒加载workbook模型
    // workbook_model_ = std::make_unique<WorkbookModel>();
    // 从source_reader_读取并解析xl/workbook.xml
}

void PackageEditor::ensureSheetModel(const SheetId& sheet) const {
    if (sheet_models_.count(sheet)) return;
    
    // 懒加载工作表模型
    auto model = std::make_unique<WorksheetModel>();
    model->name = sheet;
    sheet_models_[sheet] = std::move(model);
    
    // 如果是编辑现有文件，加载现有的工作表数据
    if (source_reader_) {
        std::string sheet_xml;
        std::string sheet_path = "xl/worksheets/" + sheet + ".xml";
        // Use extractFile instead of readEntry
        if (source_reader_->extractFile(sheet_path, sheet_xml) == archive::ZipError::Ok) {
            // TODO: 解析现有的工作表数据
        }
    }
}

void PackageEditor::ensureStylesModel() const {
    if (styles_model_) return;
    
    // 懒加载样式模型
    styles_model_ = std::make_unique<StylesModel>();
    
    // 如果是编辑现有文件，加载现有的样式
    if (source_reader_) {
        std::string styles_xml;
        // Use extractFile instead of readEntry
        if (source_reader_->extractFile("xl/styles.xml", styles_xml) == archive::ZipError::Ok) {
            // TODO: 解析现有的样式
        }
    }
}

void PackageEditor::ensureSharedStringsModel() const {
    if (sst_model_) return;
    
    // 懒加载共享字符串模型
    sst_model_ = std::make_unique<SharedStringsModel>();
    
    // 如果是编辑现有文件，加载现有的共享字符串
    if (source_reader_) {
        std::string sst_xml;
        // Use extractFile instead of readEntry
        if (source_reader_->extractFile("xl/sharedStrings.xml", sst_xml) == archive::ZipError::Ok) {
            // TODO: 解析现有的共享字符串
        }
    }
}

// 样式生成方法
std::string PackageEditor::generateStyles() const { 
    // 必须有工作簿对象才能生成styles.xml
    if (!workbook_) {
        LOG_ERROR("Cannot generate styles.xml without a Workbook object");
        return "";
    }
    
    std::ostringstream buffer;
    auto callback = [&buffer](const char* data, size_t size) {
        buffer.write(data, static_cast<std::streamsize>(size));
    };
    
    // 调用Workbook的generateStylesXML方法（现在可以访问，因为PackageEditor是友元类）
    workbook_->generateStylesXML(callback);
    return buffer.str();
}
std::string PackageEditor::generateSharedStrings() const { 
    // 必须有工作簿对象并且启用了共享字符串
    if (!workbook_) {
        LOG_ERROR("Cannot generate sharedStrings.xml without a Workbook object");
        return "";
    }
    
    if (!workbook_->getOptions().use_shared_strings) {
        // 不使用共享字符串，返回空
        return "";
    }
    
    std::ostringstream buffer;
    auto callback = [&buffer](const char* data, size_t size) {
        buffer.write(data, static_cast<std::streamsize>(size));
    };
    
    // 调用Workbook的generateSharedStringsXML方法（现在可以访问，因为PackageEditor是友元类）
    workbook_->generateSharedStringsXML(callback);
    return buffer.str();
}
std::string PackageEditor::generateContentTypes() const { return content_types_->serialize(); }
std::string PackageEditor::generateRels(const std::string& rels_path) const { 
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
    xml << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n";
    
    if (rels_path == "_rels/.rels") {
        // 根关系文件
        xml << "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\r\n";
        xml << "  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\r\n";
        xml << "  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\r\n";
    } else if (rels_path == "xl/_rels/workbook.xml.rels") {
        // 工作簿关系文件
        int rId = 1;
        
        // 工作表关系
        for (size_t i = 0; i < getSheetNames().size(); ++i) {
            xml << "  <Relationship Id=\"rId" << rId++ 
                << "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
                << "Target=\"worksheets/sheet" << (i + 1) << ".xml\"/>\r\n";
        }
        
        // 样式关系
        xml << "  <Relationship Id=\"rId" << rId++ 
            << "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
            << "Target=\"styles.xml\"/>\r\n";
        
        // 共享字符串关系（如果启用）
        if (options_.use_shared_strings) {
            xml << "  <Relationship Id=\"rId" << rId++ 
                << "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" "
                << "Target=\"sharedStrings.xml\"/>\r\n";
        }
    }
    
    xml << "</Relationships>\r\n";
    return xml.str();
}

// ========== 验证方法 ==========

bool PackageEditor::isValidSheetName(const std::string& name) {
    if (name.empty() || name.length() > 31) {
        return false;
    }
    
    // 检查禁止的字符
    static const std::string forbidden_chars = "[]\\/*?:";
    for (char c : forbidden_chars) {
        if (name.find(c) != std::string::npos) {
            return false;
        }
    }
    
    // 检查是否以单引号开头或结尾
    if (name.front() == '\'' || name.back() == '\'') {
        return false;
    }
    
    // 检查保留名称
    static const std::vector<std::string> reserved_names = {
        "History"
    };
    
    for (const auto& reserved : reserved_names) {
        if (name == reserved) {
            return false;
        }
    }
    
    return true;
}

bool PackageEditor::isValidCellRef(int row, int col) {
    return row >= 1 && row <= MAX_ROWS && col >= 1 && col <= MAX_COLS;
}

// ========== 工作表映射管理 ==========

void PackageEditor::rebuildSheetMapping() const {
    LOG_DEBUG("开始重建工作表映射...");
    sheet_name_to_id_.clear();
    sheet_id_to_name_.clear();
    
    if (!workbook_) {
        LOG_ERROR("rebuildSheetMapping: workbook_ 为空，无法重建映射");
        return;
    }
    
    auto sheet_names = workbook_->getWorksheetNames();
    LOG_DEBUG("从 Workbook 获取到 {} 个工作表", sheet_names.size());
    
    for (size_t i = 0; i < sheet_names.size(); ++i) {
        int sheet_id = static_cast<int>(i + 1);
        const auto& name = sheet_names[i];
        
        sheet_name_to_id_[name] = sheet_id;
        sheet_id_to_name_[sheet_id] = name;
        
        LOG_DEBUG("建立映射：'{}' <-> ID {}", name, sheet_id);
    }
    
    LOG_DEBUG("重建了工作表映射：{} 个工作表", sheet_names.size());
}

int PackageEditor::getSheetId(const std::string& sheet_name) const {
    // 始终重新构建映射以确保数据最新
    rebuildSheetMapping();
    
    auto it = sheet_name_to_id_.find(sheet_name);
    if (it != sheet_name_to_id_.end()) {
        LOG_DEBUG("找到工作表 '{}' 的ID：{}", sheet_name, it->second);
        return it->second;
    }
    
    LOG_ERROR("找不到工作表：'{}'", sheet_name);
    return -1;
}

std::string PackageEditor::getSheetName(int sheet_id) const {
    if (sheet_id_to_name_.empty()) {
        rebuildSheetMapping();
    }
    
    LOG_DEBUG("查找工作表ID：{}，当前映射中有 {} 个条目", sheet_id, sheet_id_to_name_.size());
    for (const auto& [id, name] : sheet_id_to_name_) {
        LOG_DEBUG("  映射：ID {} -> '{}'", id, name);
    }
    
    auto it = sheet_id_to_name_.find(sheet_id);
    if (it != sheet_id_to_name_.end()) {
        LOG_DEBUG("找到工作表ID {} 对应的名称：'{}'", sheet_id, it->second);
        return it->second;
    }
    
    LOG_ERROR("找不到工作表ID：{}，可能映射未正确建立", sheet_id);
    return "";
}

std::string PackageEditor::getSheetPath(const std::string& sheet_name) const {
    int sheet_id = getSheetId(sheet_name);
    if (sheet_id > 0) {
        return "xl/worksheets/sheet" + std::to_string(sheet_id) + ".xml";
    }
    return "";
}

std::string PackageEditor::getSheetPath(int sheet_id) const {
    if (sheet_id > 0) {
        return "xl/worksheets/sheet" + std::to_string(sheet_id) + ".xml";
    }
    return "";
}

}} // namespace fastexcel::opc
