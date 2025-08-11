#include "fastexcel/utils/ModuleLoggers.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/BatchFileWriter.hpp"
#include "fastexcel/core/CustomPropertyManager.hpp"
#include "fastexcel/core/DefinedNameManager.hpp"
#include "fastexcel/core/ExcelStructureGenerator.hpp"
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/StreamingFileWriter.hpp"
#include "fastexcel/core/StyleTransferContext.hpp"
#include "fastexcel/reader/XLSXReader.hpp"
#include "fastexcel/theme/Theme.hpp"
#include "fastexcel/theme/ThemeParser.hpp"
#include "fastexcel/utils/LogConfig.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include "fastexcel/utils/XMLUtils.hpp"
#include "fastexcel/xml/Relationships.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/xml/StyleSerializer.hpp"
#include "fastexcel/xml/UnifiedXMLGenerator.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"

#include <algorithm>
#include <ctime>
#include <fstream>
#include <functional>
#include <iomanip>
#include <sstream>

namespace fastexcel {
namespace core {

// ========== DocumentProperties 实现 ==========

DocumentProperties::DocumentProperties() {
    // 使用 TimeUtils 获取当前时间
    created_time = utils::TimeUtils::getCurrentTime();
    modified_time = created_time;
}

// ========== Workbook 实现 ==========

std::unique_ptr<Workbook> Workbook::create(const Path& path) {
    auto workbook = std::make_unique<Workbook>(path);
    
    // 🔧 新状态管理系统：创建工作簿时设置正确的状态
    workbook->file_source_ = FileSource::NEW_FILE;
    workbook->transitionToState(WorkbookState::CREATING, "Workbook::create()");
    
    // 🔧 关键修复：对于 create() 创建的工作簿，强制设置为新文件
    // 因为我们要完全重写目标文件，无论它是否已存在
    if (workbook->dirty_manager_) {
        workbook->dirty_manager_->setIsNewFile(true);
    }
    
    // 🎯 API修复：自动打开工作簿，返回可直接使用的对象
    if (!workbook->open()) {
        CORE_ERROR("Failed to open workbook after creation: {}", path.string());
        return nullptr;
    }
    
    return workbook;
}

Workbook::Workbook(const Path& path) : filename_(path.string()) {
    // 检查是否为内存模式（任何以::memory::开头的路径）
    if (path.string().find("::memory::") == 0) {
        // 内存模式：不创建FileManager，保持纯内存操作
        file_manager_ = nullptr;
        CORE_DEBUG("Created workbook in memory mode: {}", filename_);
    } else {
        // 文件模式：创建FileManager处理文件操作
        file_manager_ = std::make_unique<archive::FileManager>(path);
    }
    
    format_repo_ = std::make_unique<FormatRepository>();
    // 初始化共享字符串表
    shared_string_table_ = std::make_unique<SharedStringTable>();
    
    // 初始化管理器
    custom_property_manager_ = std::make_unique<CustomPropertyManager>();
    defined_name_manager_ = std::make_unique<DefinedNameManager>();
    
    // 初始化智能脏数据管理器
    dirty_manager_ = std::make_unique<DirtyManager>();
    // 🔧 关键修复：对于 create() 创建的工作簿，无论目标文件是否存在都视为新文件
    // 这里暂时保持原逻辑，在 create() 方法中会重新设置
    dirty_manager_->setIsNewFile(!path.exists()); // 如果文件不存在，则是新文件
    
    // 设置默认文档属性
    doc_properties_.author = "FastExcel";
    doc_properties_.company = "FastExcel Library";
}

Workbook::~Workbook() {
    close();
}

// ========== 文件操作 ==========

bool Workbook::open() {
    // 内存模式无需文件操作
    if (!file_manager_) {
        CORE_DEBUG("Memory workbook opened: {}", filename_);
        return true;
    }
    
    // 文件模式需要打开FileManager
    bool success = file_manager_->open(true);
    if (success) {
        CORE_INFO("Workbook opened: {}", filename_);
    }
    
    return success;
}

bool Workbook::save() {
    // 运行时检查：只读模式不能保存
    ensureEditable("save");
    
    try {
        // 使用 TimeUtils 更新修改时间
        doc_properties_.modified_time = utils::TimeUtils::getCurrentTime();
        
        // 设置ZIP压缩级别
        if (file_manager_->isOpen()) {
            if (!file_manager_->setCompressionLevel(options_.compression_level)) {
                CORE_WARN("Failed to set compression level to {}", options_.compression_level);
            } else {
                FASTEXCEL_LOG_ZIP_DEBUG("Set ZIP compression level to {}", options_.compression_level);
            }
        }
        
        // 🔧 修复SharedStrings生成逻辑：移除手动收集，依赖工作表XML生成时自动添加
        // 清空共享字符串列表，让工作表XML生成时自动填充
        if (options_.use_shared_strings) {
            CORE_DEBUG("SharedStrings enabled - SST will be populated during worksheet XML generation");
            if (shared_string_table_) shared_string_table_->clear();
        } else {
            CORE_DEBUG("SharedStrings disabled for performance");
            if (shared_string_table_) shared_string_table_->clear();
        }
        
        // 编辑模式下，先将原包中未被我们生成的条目拷贝过来（绘图、图片、打印设置等）
        if (isPassThroughEditMode() && !original_package_path_.empty() && file_manager_ && file_manager_->isOpen()) {
            // 我们将跳过这些前缀（由生成逻辑负责写入/覆盖）
            // 透传阶段：不跳过任何前缀，先复制全部条目；后续生成阶段会覆盖我们需要更新的部件
            std::vector<std::string> skip_prefixes = { };
            file_manager_->copyFromExistingPackage(core::Path(original_package_path_), skip_prefixes);
        }

        // 生成Excel文件结构（会覆盖我们管理的核心部件）
        if (!generateExcelStructure()) {
            CORE_ERROR("Failed to generate Excel structure");
            return false;
        }
        
        CORE_INFO("Workbook saved successfully: {}", filename_);
        return true;
    } catch (const std::exception& e) {
        CORE_ERROR("Failed to save workbook: {}", e.what());
        return false;
    }
}

bool Workbook::saveAs(const std::string& filename) {
    // 运行时检查：只读模式不能保存
    ensureEditable("saveAs");
    
    std::string old_filename = filename_;
    std::string original_source = original_package_path_;
    bool was_from_existing = (file_source_ == FileSource::EXISTING_FILE);

    // 检查是否保存到同一个文件
    bool is_same_file = (filename == old_filename) || (filename == original_source);
    
    if (is_same_file && was_from_existing && !original_source.empty()) {
        // 如果保存到同一个文件，需要先复制原文件到临时位置
        CORE_INFO("Saving to same file, creating temporary backup for resource preservation");
        
        // 创建临时文件路径
        std::string temp_backup = original_source + ".tmp_backup";
        core::Path source_path(original_source);
        core::Path temp_path(temp_backup);
        
        // 复制原文件到临时位置
        try {
            if (temp_path.exists()) {
                temp_path.remove();
            }
            source_path.copyTo(temp_path);
            original_package_path_ = temp_backup;  // 更新源路径为临时文件
            CORE_DEBUG("Created temporary backup: {}", temp_backup);
        } catch (const std::exception& e) {
            CORE_ERROR("Failed to create temporary backup: {}", e.what());
            return false;
        }
    }

    filename_ = filename;
    
    // 重新创建文件管理器
    file_manager_ = std::make_unique<archive::FileManager>(core::Path(filename));
    
    if (!file_manager_->open(true)) {
        // 恢复原文件名
        filename_ = old_filename;
        file_manager_ = std::make_unique<archive::FileManager>(core::Path(old_filename));
        
        // 如果创建了临时文件，删除它
        if (is_same_file && original_package_path_.find(".tmp_backup") != std::string::npos) {
            core::Path temp_path(original_package_path_);
            if (temp_path.exists()) {
                temp_path.remove();
            }
            original_package_path_ = original_source;  // 恢复原路径
        }
        return false;
    }

    // 在另存为场景下，如果当前工作簿是从现有包打开的，那么保留 original_package_path_ 用于拷贝未修改部件
    // file_source_ 状态保持不变，已经在之前设置好了
    // original_package_path_ 已经在上面设置好了（可能是临时文件或原始文件）
    
    bool save_result = save();
    
    // 清理临时文件（如果有）
    if (is_same_file && original_package_path_.find(".tmp_backup") != std::string::npos) {
        core::Path temp_path(original_package_path_);
        if (temp_path.exists()) {
            temp_path.remove();
            CORE_DEBUG("Removed temporary backup: {}", original_package_path_);
        }
        original_package_path_ = original_source;  // 恢复原路径
    }
    
    return save_result;
}

bool Workbook::isOpen() const {
    // 检查工作簿是否处于可用状态（EDITING、READING或CREATING模式）
    return state_ == WorkbookState::EDITING || 
           state_ == WorkbookState::READING || 
           state_ == WorkbookState::CREATING;
}

bool Workbook::close() {
    // 内存模式只需要重置状态
    if (!file_manager_) {
        CORE_DEBUG("Memory workbook closed: {}", filename_);
    } else {
        // 文件模式需要关闭FileManager
        file_manager_->close();
        CORE_INFO("Workbook closed: {}", filename_);
    }
    return true;
}

// ========== 工作表管理 ==========

std::shared_ptr<Worksheet> Workbook::addWorksheet(const std::string& name) {
    // 运行时检查：只读模式不能添加工作表
    ensureEditable("addWorksheet");
    
    std::string sheet_name;
    if (name.empty()) {
        sheet_name = generateUniqueSheetName("Sheet1");
    } else {
        // 检查名称是否已存在，如果存在则生成唯一名称
        if (getWorksheet(name) != nullptr) {
            sheet_name = generateUniqueSheetName(name);
        } else {
            sheet_name = name;
        }
    }
    
    if (!validateSheetName(sheet_name)) {
        CORE_ERROR("Invalid sheet name: {}", sheet_name);
        return nullptr;
    }
    
    auto worksheet = std::make_shared<Worksheet>(sheet_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    worksheets_.push_back(worksheet);
    
    // 关键修复：如果这是第一个工作表，自动设置为激活状态
    if (worksheets_.size() == 1) {
        worksheet->setTabSelected(true);
        CORE_DEBUG("Added worksheet: {} (activated as first sheet)", sheet_name);
    } else {
        CORE_DEBUG("Added worksheet: {}", sheet_name);
    }
    
    return worksheet;
}

std::shared_ptr<Worksheet> Workbook::insertWorksheet(size_t index, const std::string& name) {
    // 运行时检查：只读模式不能插入工作表
    ensureEditable("insertWorksheet");
    
    if (index > worksheets_.size()) {
        index = worksheets_.size();
    }
    
    std::string sheet_name = name.empty() ? generateUniqueSheetName("Sheet1") : name;
    
    if (!validateSheetName(sheet_name)) {
        CORE_ERROR("Invalid sheet name: {}", sheet_name);
        return nullptr;
    }
    
    auto worksheet = std::make_shared<Worksheet>(sheet_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    worksheets_.insert(worksheets_.begin() + index, worksheet);
    
    CORE_DEBUG("Inserted worksheet: {} at index {}", sheet_name, index);
    return worksheet;
}

bool Workbook::removeWorksheet(const std::string& name) {
    // 运行时检查：只读模式不能删除工作表
    ensureEditable("removeWorksheet");
    
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        worksheets_.erase(it);
        CORE_DEBUG("Removed worksheet: {}", name);
        return true;
    }
    
    return false;
}

bool Workbook::removeWorksheet(size_t index) {
    // 运行时检查：只读模式不能删除工作表
    ensureEditable("removeWorksheet");
    
    if (index < worksheets_.size()) {
        std::string name = worksheets_[index]->getName();
        worksheets_.erase(worksheets_.begin() + index);
        CORE_DEBUG("Removed worksheet: {} at index {}", name, index);
        return true;
    }
    
    return false;
}

std::shared_ptr<Worksheet> Workbook::getWorksheet(const std::string& name) {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(), 
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        return *it;
    }
    
    return nullptr;
}

std::shared_ptr<Worksheet> Workbook::getWorksheet(size_t index) {
    if (index < worksheets_.size()) {
        return worksheets_[index];
    }
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::getWorksheet(const std::string& name) const {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        return *it;
    }
    
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::getWorksheet(size_t index) const {
    if (index < worksheets_.size()) {
        return worksheets_[index];
    }
    return nullptr;
}

std::vector<std::string> Workbook::getWorksheetNames() const {
    std::vector<std::string> names;
    names.reserve(worksheets_.size());
    
    for (const auto& worksheet : worksheets_) {
        names.push_back(worksheet->getName());
    }
    
    return names;
}

bool Workbook::renameWorksheet(const std::string& old_name, const std::string& new_name) {
    auto worksheet = getWorksheet(old_name);
    if (!worksheet) {
        return false;
    }
    
    if (!validateSheetName(new_name)) {
        return false;
    }
    
    worksheet->setName(new_name);
    CORE_DEBUG("Renamed worksheet: {} -> {}", old_name, new_name);
    return true;
}

bool Workbook::moveWorksheet(size_t from_index, size_t to_index) {
    if (from_index >= worksheets_.size() || to_index >= worksheets_.size()) {
        return false;
    }
    
    if (from_index == to_index) {
        return true;
    }
    
    auto worksheet = worksheets_[from_index];
    worksheets_.erase(worksheets_.begin() + from_index);
    
    if (to_index > from_index) {
        to_index--;
    }
    
    worksheets_.insert(worksheets_.begin() + to_index, worksheet);
    
    CORE_DEBUG("Moved worksheet from index {} to {}", from_index, to_index);
    return true;
}

std::shared_ptr<Worksheet> Workbook::copyWorksheet(const std::string& source_name, const std::string& new_name) {
    auto source_worksheet = getWorksheet(source_name);
    if (!source_worksheet) {
        return nullptr;
    }
    
    if (!validateSheetName(new_name)) {
        return nullptr;
    }
    
    // 创建新工作表
    auto new_worksheet = std::make_shared<Worksheet>(new_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    
    // 这里应该实现深拷贝逻辑
    // 简化版本，实际需要复制所有单元格、格式、设置等
    
    worksheets_.push_back(new_worksheet);
    
    CORE_DEBUG("Copied worksheet: {} -> {}", source_name, new_name);
    return new_worksheet;
}

void Workbook::setActiveWorksheet(size_t index) {
    // 取消所有工作表的选中状态
    for (auto& worksheet : worksheets_) {
        worksheet->setTabSelected(false);
    }
    
    // 设置指定工作表为活动状态
    if (index < worksheets_.size()) {
        worksheets_[index]->setTabSelected(true);
    }
}

// ========== 样式管理 ==========

int Workbook::addStyle(const FormatDescriptor& style) {
    return format_repo_->addFormat(style);
}

int Workbook::addStyle(const StyleBuilder& builder) {
    auto format = builder.build();
    return format_repo_->addFormat(format);
}

std::shared_ptr<const FormatDescriptor> Workbook::getStyle(int style_id) const {
    // 从格式仓储中根据ID获取格式描述符
    return format_repo_->getFormat(style_id);
}

int Workbook::getDefaultStyleId() const {
    return format_repo_->getDefaultFormatId();
}

bool Workbook::isValidStyleId(int style_id) const {
    return format_repo_->isValidFormatId(style_id);
}

const FormatRepository& Workbook::getStyleRepository() const {
    return *format_repo_;
}

void Workbook::setThemeXML(const std::string& theme_xml) {
    theme_xml_ = theme_xml;
    theme_dirty_ = true; // 外部显式设置主题XML视为编辑
    CORE_DEBUG("设置自定义主题XML ({} 字节)", theme_xml_.size());
    // 尝试解析为结构化主题对象
    if (!theme_xml_.empty()) {
        auto parsed = theme::ThemeParser::parseFromXML(theme_xml_);
        if (parsed) {
            theme_ = std::move(parsed);
            CORE_DEBUG("主题XML已解析为对象: {}", theme_->getName());
        } else {
            CORE_WARN("主题XML解析失败，保留原始XML");
        }
    }
}

const std::string& Workbook::getThemeXML() const {
    return theme_xml_;
}

void Workbook::setOriginalThemeXML(const std::string& theme_xml) {
    theme_xml_original_ = theme_xml;
    CORE_DEBUG("保存原始主题XML ({} 字节)", theme_xml_original_.size());
    // 同步解析一次，便于后续编辑
    if (!theme_xml_original_.empty()) {
        auto parsed = theme::ThemeParser::parseFromXML(theme_xml_original_);
        if (parsed) {
            theme_ = std::move(parsed);
            CORE_DEBUG("原始主题XML已解析为对象: {}", theme_->getName());
        }
    }
}

void Workbook::setTheme(const theme::Theme& theme) {
    theme_ = std::make_unique<theme::Theme>(theme);
    // 同步XML缓存
    theme_xml_ = theme_->toXML();
    theme_dirty_ = true;
}

void Workbook::setThemeName(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->setName(name);
    theme_xml_.clear(); // 让生成时重新序列化
    theme_dirty_ = true;
}

void Workbook::setThemeColor(theme::ThemeColorScheme::ColorType type, const core::Color& color) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->colors().setColor(type, color);
    theme_xml_.clear();
    theme_dirty_ = true;
}

bool Workbook::setThemeColorByName(const std::string& name, const core::Color& color) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    bool ok = theme_->colors().setColorByName(name, color);
    if (ok) { theme_xml_.clear(); theme_dirty_ = true; }
    return ok;
}

void Workbook::setThemeMajorFontLatin(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMajorFontLatin(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMajorFontEastAsia(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMajorFontEastAsia(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMajorFontComplex(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMajorFontComplex(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMinorFontLatin(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMinorFontLatin(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMinorFontEastAsia(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMinorFontEastAsia(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMinorFontComplex(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMinorFontComplex(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

// ========== 自定义属性 ==========

void Workbook::setCustomProperty(const std::string& name, const std::string& value) {
    custom_property_manager_->setProperty(name, value);
}

void Workbook::setCustomProperty(const std::string& name, double value) {
    custom_property_manager_->setProperty(name, value);
}

void Workbook::setCustomProperty(const std::string& name, bool value) {
    custom_property_manager_->setProperty(name, value);
}

std::string Workbook::getCustomProperty(const std::string& name) const {
    return custom_property_manager_->getProperty(name);
}

bool Workbook::removeCustomProperty(const std::string& name) {
    return custom_property_manager_->removeProperty(name);
}

std::unordered_map<std::string, std::string> Workbook::getCustomProperties() const {
    return custom_property_manager_->getAllProperties();
}

// ========== 定义名称 ==========

void Workbook::defineName(const std::string& name, const std::string& formula, const std::string& scope) {
    defined_name_manager_->define(name, formula, scope);
}

std::string Workbook::getDefinedName(const std::string& name, const std::string& scope) const {
    return defined_name_manager_->get(name, scope);
}

bool Workbook::removeDefinedName(const std::string& name, const std::string& scope) {
    return defined_name_manager_->remove(name, scope);
}

// ========== VBA项目 ==========

bool Workbook::addVbaProject(const std::string& vba_project_path) {
    // 检查文件是否存在
    std::ifstream file(vba_project_path, std::ios::binary);
    if (!file.is_open()) {
        CORE_ERROR("VBA project file not found: {}", vba_project_path);
        return false;
    }
    
    vba_project_path_ = vba_project_path;
    has_vba_ = true;
    
    CORE_INFO("Added VBA project: {}", vba_project_path);
    return true;
}

// ========== 工作簿保护 ==========

void Workbook::protect(const std::string& password, bool lock_structure, bool lock_windows) {
    protected_ = true;
    protection_password_ = password;
    lock_structure_ = lock_structure;
    lock_windows_ = lock_windows;
}

void Workbook::unprotect() {
    protected_ = false;
    protection_password_.clear();
    lock_structure_ = false;
    lock_windows_ = false;
}

// ========== 工作簿选项 ==========

void Workbook::setCalcOptions(bool calc_on_load, bool full_calc_on_load) {
    options_.calc_on_load = calc_on_load;
    options_.full_calc_on_load = full_calc_on_load;
}

// ========== 生成控制判定（使用DirtyManager智能管理） ==========

bool Workbook::shouldGenerateContentTypes() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("[Content_Types].xml");
}

bool Workbook::shouldGenerateRootRels() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("_rels/.rels");
}

bool Workbook::shouldGenerateWorkbookCore() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("xl/workbook.xml");
}

bool Workbook::shouldGenerateStyles() const {
    // 始终生成样式文件，保证包内引用一致性：
    // - workbook.xml 和 [Content_Types].xml 总是包含对 xl/styles.xml 的引用
    // - 如不生成，将导致包缺少被引用的部件，Excel 打开会提示修复
    // 样式文件很小，生成最小可用样式的成本可以忽略
    return true;
}

bool Workbook::shouldGenerateTheme() const {
    // 🔧 关键修复：只有在确实有主题内容时才生成主题文件
    // 避免请求生成主题但ThemeGenerator找不到内容的问题
    if (!theme_xml_.empty() || !theme_xml_original_.empty() || theme_) {
        return true;
    }
    return false; // 没有主题内容，不生成主题文件
}

bool Workbook::shouldGenerateSharedStrings() const {
    CORE_DEBUG("shouldGenerateSharedStrings() called - analyzing conditions");
    
    if (!options_.use_shared_strings) {
        CORE_DEBUG("SharedStrings generation disabled by options_.use_shared_strings = false");
        return false; // 未启用SST
    }
    CORE_DEBUG("options_.use_shared_strings = true, SharedStrings enabled");
    
    if (!dirty_manager_) {
        CORE_DEBUG("No dirty manager, SharedStrings generation enabled (default true)");
        return true;
    }
    CORE_DEBUG("DirtyManager exists, checking shouldUpdate for xl/sharedStrings.xml");
    
    bool should_update = dirty_manager_->shouldUpdate("xl/sharedStrings.xml");
    CORE_DEBUG("DirtyManager shouldUpdate for SharedStrings: {}", should_update);
    
    // 🔧 关键修复：如果SharedStringTable有内容但DirtyManager说不需要更新，强制生成
    if (shared_string_table_) {
        size_t string_count = shared_string_table_->getStringCount();
        CORE_DEBUG("SharedStringTable contains {} strings", string_count);
        
        if (string_count > 0 && !should_update) {
            CORE_DEBUG("🔧 FORCE GENERATION: SharedStringTable has {} strings but DirtyManager says no update needed", string_count);
            CORE_DEBUG("🔧 This happens when target file exists but we're creating new content with strings");
            CORE_DEBUG("🔧 Forcing SharedStrings generation to avoid missing sharedStrings.xml");
            return true; // 强制生成
        }
    } else {
        CORE_DEBUG("SharedStringTable is null");
    }
    
    return should_update;
}

bool Workbook::shouldGenerateDocPropsCore() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("docProps/core.xml");
}

bool Workbook::shouldGenerateDocPropsApp() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("docProps/app.xml");
}

bool Workbook::shouldGenerateDocPropsCustom() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("docProps/custom.xml");
}

bool Workbook::shouldGenerateSheet(size_t index) const {
    if (!dirty_manager_) return true;
    std::string sheetPart = "xl/worksheets/sheet" + std::to_string(index + 1) + ".xml";
    return dirty_manager_->shouldUpdate(sheetPart);
}

bool Workbook::shouldGenerateSheetRels(size_t index) const {
    if (!dirty_manager_) return true;
    std::string sheetRelsPart = "xl/worksheets/_rels/sheet" + std::to_string(index + 1) + ".xml.rels";
    return dirty_manager_->shouldUpdate(sheetRelsPart);
}

// ========== 共享字符串管理 ==========

int Workbook::addSharedString(const std::string& str) {
    if (!shared_string_table_) shared_string_table_ = std::make_unique<SharedStringTable>();
    return static_cast<int>(shared_string_table_->addString(str));
}

int Workbook::addSharedStringWithIndex(const std::string& str, int original_index) {
    if (!shared_string_table_) shared_string_table_ = std::make_unique<SharedStringTable>();
    return static_cast<int>(shared_string_table_->addStringWithId(str, original_index));
}

int Workbook::getSharedStringIndex(const std::string& str) const {
    if (!shared_string_table_) return -1;
    return static_cast<int>(shared_string_table_->getStringId(str));
}

const SharedStringTable* Workbook::getSharedStringTable() const {
    return shared_string_table_.get();
}

// ========== 内部方法 ==========

bool Workbook::generateExcelStructure() {
    // 智能选择生成模式：根据数据量和内存使用情况自动决定
    size_t estimated_memory = estimateMemoryUsage();
    size_t total_cells = getTotalCellCount();
    
    bool use_streaming = false;
    
    // 新的决策逻辑：基于WorkbookMode
    switch (options_.mode) {
        case WorkbookMode::AUTO:
            // 自动模式：根据数据量智能选择
            if (total_cells > options_.auto_mode_cell_threshold ||
                estimated_memory > options_.auto_mode_memory_threshold) {
                use_streaming = true;
                CORE_INFO("Auto-selected streaming mode: {} cells, {}MB estimated memory (thresholds: {} cells, {}MB)",
                        total_cells, estimated_memory / (1024*1024),
                        options_.auto_mode_cell_threshold, options_.auto_mode_memory_threshold / (1024*1024));
            } else {
                use_streaming = false;
                CORE_INFO("Auto-selected batch mode: {} cells, {}MB estimated memory (thresholds: {} cells, {}MB)",
                        total_cells, estimated_memory / (1024*1024),
                        options_.auto_mode_cell_threshold, options_.auto_mode_memory_threshold / (1024*1024));
            }
            break;
            
        case WorkbookMode::BATCH:
            // 强制批量模式
            use_streaming = false;
            CORE_INFO("Using forced batch mode: {} cells, {}MB estimated memory",
                    total_cells, estimated_memory / (1024*1024));
            break;
            
        case WorkbookMode::STREAMING:
            // 强制流式模式
            use_streaming = true;
            CORE_INFO("Using forced streaming mode: {} cells, {}MB estimated memory",
                    total_cells, estimated_memory / (1024*1024));
            break;
    }
    
    // 如果设置了constant_memory，强制使用流式模式
    if (options_.constant_memory) {
        use_streaming = true;
        CORE_INFO("Constant memory mode enabled, forcing streaming mode");
    }
    
    return generateWithGenerator(use_streaming);
}




// 主题写出逻辑已迁移至 XML 层（ThemeGenerator），此处不再直接输出

// ========== 格式管理内部方法 ==========


// ========== 辅助函数 ==========

std::string Workbook::generateUniqueSheetName(const std::string& base_name) const {
    // 如果base_name不存在，直接返回
    if (getWorksheet(base_name) == nullptr) {
        return base_name;
    }
    
    // 如果base_name是"Sheet1"，从"Sheet2"开始尝试
    if (base_name == "Sheet1") {
        int counter = 2;
        std::string name = "Sheet" + std::to_string(counter);
        while (getWorksheet(name) != nullptr) {
            name = "Sheet" + std::to_string(++counter);
        }
        return name;
    }
    
    // 对于其他base_name，添加数字后缀
    int suffix_counter = 1;
    std::string name = base_name + std::to_string(suffix_counter);
    while (getWorksheet(name) != nullptr) {
        name = base_name + std::to_string(++suffix_counter);
    }
    
    return name;
}

bool Workbook::validateSheetName(const std::string& name) const {
    // 检查长度
    if (name.empty() || name.length() > 31) {
        return false;
    }
    
    // 检查非法字符
    const std::string invalid_chars = ":\\/?*[]";
    if (name.find_first_of(invalid_chars) != std::string::npos) {
        return false;
    }
    
    // 检查是否以单引号开头或结尾
    if (name.front() == '\'' || name.back() == '\'') {
        return false;
    }
    
    // 不检查是否已存在，因为这个方法也被用于验证新名称
    // 重复名称的检查应该在调用方处理
    
    return true;
}

void Workbook::collectSharedStrings() {
    if (!shared_string_table_) {
        shared_string_table_ = std::make_unique<SharedStringTable>();
    } else {
        shared_string_table_->clear();
    }
    
    for (const auto& worksheet : worksheets_) {
        // 这里需要访问工作表的单元格来收集字符串
        // 简化版本，实际实现需要遍历所有字符串单元格
        auto [max_row, max_col] = worksheet->getUsedRange();
        
        for (int row = 0; row <= max_row; ++row) {
            for (int col = 0; col <= max_col; ++col) {
                if (worksheet->hasCellAt(row, col)) {
                    const auto& cell = worksheet->getCell(row, col);
                    if (cell.isString()) {
                        addSharedString(cell.getStringValue());
                    }
                }
            }
        }
    }
}


std::string Workbook::getWorksheetPath(int sheet_id) const {
    return "xl/worksheets/sheet" + std::to_string(sheet_id) + ".xml";
}

std::string Workbook::getWorksheetRelPath(int sheet_id) const {
    return "worksheets/sheet" + std::to_string(sheet_id) + ".xml";
}


void Workbook::setHighPerformanceMode(bool enable) {
    if (enable) {
        CORE_INFO("Enabling ultra high performance mode (beyond defaults)");
        
        // 进一步优化：使用无压缩模式排除压缩算法影响
        options_.compression_level = 0;  // 无压缩
        
        // 更大的缓冲区
        options_.row_buffer_size = 10000;
        options_.xml_buffer_size = 8 * 1024 * 1024;  // 8MB
        
        // 使用AUTO模式，让系统根据数据量自动选择
        options_.mode = WorkbookMode::AUTO;
        options_.use_shared_strings = true;
        
        // 调整自动模式阈值，更倾向于使用批量模式以获得更好的性能
        options_.auto_mode_cell_threshold = 2000000;  // 200万单元格
        options_.auto_mode_memory_threshold = 200 * 1024 * 1024;  // 200MB
        
        CORE_INFO("Ultra high performance mode configured: Mode=AUTO, Compression=OFF, RowBuffer={}, XMLBuffer={}MB",
                options_.row_buffer_size, options_.xml_buffer_size / (1024*1024));
    } else {
        CORE_INFO("Using standard high performance mode (default settings)");
        
        // 恢复到默认的高性能设置
        options_.mode = WorkbookMode::AUTO;           // 默认自动模式
        options_.use_shared_strings = true;           // 默认启用以匹配Excel格式
        options_.row_buffer_size = 5000;              // 默认较大缓冲
        options_.compression_level = 6;               // 恢复到默认的中等压缩
        options_.xml_buffer_size = 4 * 1024 * 1024;  // 默认4MB
        
        // 恢复默认阈值
        options_.auto_mode_cell_threshold = 1000000;     // 100万单元格
        options_.auto_mode_memory_threshold = 100 * 1024 * 1024; // 100MB
    }
}




// ========== 工作簿编辑功能实现 ==========

std::unique_ptr<Workbook> Workbook::open(const Path& path) {
    try {
        // 使用Path的内置文件检查
        if (!path.exists()) {
            CORE_ERROR("File not found for editing: {}", path.string());
            return nullptr;
        }
        
        // 使用XLSXReader读取现有文件
        reader::XLSXReader reader(path);
        auto result = reader.open();
        if (result != core::ErrorCode::Ok) {
            CORE_ERROR("Failed to open XLSX file for reading: {}, error code: {}", path.string(), static_cast<int>(result));
            return nullptr;
        }
        
        // 加载工作簿
        std::unique_ptr<core::Workbook> loaded_workbook;
        result = reader.loadWorkbook(loaded_workbook);
        reader.close();
        
        if (result != core::ErrorCode::Ok) {
            CORE_ERROR("Failed to load workbook from file: {}, error code: {}", path.string(), static_cast<int>(result));
            return nullptr;
        }
        
        // 标记来源以便保存时进行未修改部件的保真写回
        if (loaded_workbook) {
            loaded_workbook->transitionToState(WorkbookState::EDITING, "openForEditing()");
            loaded_workbook->original_package_path_ = path.string();
            // 设置为编辑模式（保持向后兼容）
            loaded_workbook->file_source_ = FileSource::EXISTING_FILE;
            
            // 🎯 API修复：为保存功能准备FileManager
            if (!loaded_workbook->open()) {
                CORE_ERROR("Failed to prepare FileManager for workbook: {}", path.string());
                return nullptr;
            }
        }
        
        CORE_INFO("Successfully loaded workbook for editing: {}", path.string());
        return loaded_workbook;
        
    } catch (const std::exception& e) {
        CORE_ERROR("Exception while loading workbook for editing: {}, error: {}", path.string(), e.what());
        return nullptr;
    }
}

// ========== 新的语义化API实现 ==========

std::unique_ptr<Workbook> Workbook::openForReading(const Path& path) {
    try {
        if (!path.exists()) {
            CORE_ERROR("File not found for reading: {}", path.string());
            return nullptr;
        }
        
        // 使用XLSXReader读取现有文件
        reader::XLSXReader reader(path);
        auto result = reader.open();
        if (result != core::ErrorCode::Ok) {
            CORE_ERROR("Failed to open XLSX file for reading: {}, error code: {}", path.string(), static_cast<int>(result));
            return nullptr;
        }
        
        // 加载工作簿
        std::unique_ptr<core::Workbook> loaded_workbook;
        result = reader.loadWorkbook(loaded_workbook);
        reader.close();
        
        if (result != core::ErrorCode::Ok) {
            CORE_ERROR("Failed to load workbook from file: {}, error code: {}", path.string(), static_cast<int>(result));
            return nullptr;
        }
        
        // 设置为只读模式
        if (loaded_workbook) {
            loaded_workbook->file_source_ = FileSource::EXISTING_FILE;
            loaded_workbook->transitionToState(WorkbookState::READING, "openForReading()");
            loaded_workbook->original_package_path_ = path.string();
            
            // 🎯 API修复：为保存功能准备FileManager（即使是只读模式，也可能需要另存为）
            if (!loaded_workbook->open()) {
                CORE_ERROR("Failed to prepare FileManager for workbook: {}", path.string());
                return nullptr;
            }
            
            // 只读模式优化：后续可增加更细粒度的追踪开关，这里不额外操作
        }
        
        CORE_INFO("Successfully loaded workbook for reading: {}", path.string());
        return loaded_workbook;
        
    } catch (const std::exception& e) {
        CORE_ERROR("Exception while loading workbook for reading: {}, error: {}", path.string(), e.what());
        return nullptr;
    }
}

std::unique_ptr<Workbook> Workbook::openForEditing(const Path& path) {
    // 编辑模式就是原有的open方法实现
    try {
        if (!path.exists()) {
            CORE_ERROR("File not found for editing: {}", path.string());
            return nullptr;
        }
        
        // 使用XLSXReader读取现有文件
        reader::XLSXReader reader(path);
        auto result = reader.open();
        if (result != core::ErrorCode::Ok) {
            CORE_ERROR("Failed to open XLSX file for editing: {}, error code: {}", path.string(), static_cast<int>(result));
            return nullptr;
        }
        
        // 加载工作簿
        std::unique_ptr<core::Workbook> loaded_workbook;
        result = reader.loadWorkbook(loaded_workbook);
        reader.close();
        
        if (result != core::ErrorCode::Ok) {
            CORE_ERROR("Failed to load workbook from file: {}, error code: {}", path.string(), static_cast<int>(result));
            return nullptr;
        }
        
        // 设置为编辑模式
        if (loaded_workbook) {
            loaded_workbook->file_source_ = FileSource::EXISTING_FILE;
            loaded_workbook->transitionToState(WorkbookState::EDITING, "openForEditing()");
            loaded_workbook->original_package_path_ = path.string();
            
            // 🎯 API修复：为保存功能准备FileManager
            if (!loaded_workbook->open()) {
                CORE_ERROR("Failed to prepare FileManager for workbook: {}", path.string());
                return nullptr;
            }
        }
        
        CORE_INFO("Successfully loaded workbook for editing: {}", path.string());
        return loaded_workbook;
        
    } catch (const std::exception& e) {
        CORE_ERROR("Exception while loading workbook for editing: {}, error: {}", path.string(), e.what());
        return nullptr;
    }
}

bool Workbook::refresh() {
    
    try {
        // 保存当前状态
        std::string current_filename = filename_;
        
        // 关闭当前工作簿
        close();
        
        // 重新加载
        Path current_path(current_filename);
        auto refreshed_workbook = open(current_path);
        if (!refreshed_workbook) {
            CORE_ERROR("Failed to refresh workbook: {}", current_filename);
            return false;
        }
        
        // 替换当前内容
        worksheets_ = std::move(refreshed_workbook->worksheets_);
        format_repo_ = std::move(refreshed_workbook->format_repo_);
        doc_properties_ = refreshed_workbook->doc_properties_;
        custom_property_manager_ = std::move(refreshed_workbook->custom_property_manager_);
        defined_name_manager_ = std::move(refreshed_workbook->defined_name_manager_);
        
        // 重新打开工作簿
        open();
        
        CORE_INFO("Successfully refreshed workbook: {}", current_filename);
        return true;
        
    } catch (const std::exception& e) {
        CORE_ERROR("Exception during workbook refresh: {}", e.what());
        return false;
    }
}

bool Workbook::mergeWorkbook(const std::unique_ptr<Workbook>& other_workbook, const MergeOptions& options) {
    if (!other_workbook) {
        CORE_ERROR("Cannot merge: other workbook is null");
        return false;
    }
    
    
    try {
        int merged_count = 0;
        
        // 合并工作表
        if (options.merge_worksheets) {
            for (const auto& other_worksheet : other_workbook->worksheets_) {
                std::string new_name = options.name_prefix + other_worksheet->getName();
                
                // 检查名称冲突
                if (getWorksheet(new_name) != nullptr) {
                    if (options.overwrite_existing) {
                        removeWorksheet(new_name);
                        CORE_INFO("Removed existing worksheet for merge: {}", new_name);
                    } else {
                        new_name = generateUniqueSheetName(new_name);
                        CORE_INFO("Generated unique name for merge: {}", new_name);
                    }
                }
                
                // 创建新工作表并复制内容
                auto new_worksheet = addWorksheet(new_name);
                if (new_worksheet) {
                    // 这里需要实现深拷贝逻辑
                    // 简化版本：复制基本属性
                    merged_count++;
                    CORE_DEBUG("Merged worksheet: {} -> {}", other_worksheet->getName(), new_name);
                }
            }
        }
        
        // 合并格式
        if (options.merge_formats) {
            // 将其他工作簿的格式仓储合并到当前格式仓储
            // 遍历其他工作簿的所有格式并添加到当前仓储中（自动去重）
            for (const auto& format_item : *other_workbook->format_repo_) {
                format_repo_->addFormat(*format_item.format);
            }
            CORE_DEBUG("Merged formats from other workbook");
        }
        
        // 合并文档属性
        if (options.merge_properties) {
            if (!other_workbook->doc_properties_.title.empty()) {
                doc_properties_.title = other_workbook->doc_properties_.title;
            }
            if (!other_workbook->doc_properties_.author.empty()) {
                doc_properties_.author = other_workbook->doc_properties_.author;
            }
            if (!other_workbook->doc_properties_.subject.empty()) {
                doc_properties_.subject = other_workbook->doc_properties_.subject;
            }
            if (!other_workbook->doc_properties_.company.empty()) {
                doc_properties_.company = other_workbook->doc_properties_.company;
            }
            
            // 合并自定义属性
            for (const auto& prop : other_workbook->custom_property_manager_->getAllDetailedProperties()) {
                setCustomProperty(prop.name, prop.value);
            }
            
            CORE_DEBUG("Merged document properties");
        }
        
        CORE_INFO("Successfully merged workbook: {} worksheets, {} formats",
                merged_count, other_workbook->format_repo_->getFormatCount());
        return true;
        
    } catch (const std::exception& e) {
        CORE_ERROR("Exception during workbook merge: {}", e.what());
        return false;
    }
}

bool Workbook::exportWorksheets(const std::vector<std::string>& worksheet_names, const std::string& output_filename) {
    if (worksheet_names.empty()) {
        CORE_ERROR("No worksheets specified for export");
        return false;
    }
    
    try {
        // 创建新工作簿
        auto export_workbook = create(Path(output_filename));
        if (!export_workbook->open()) {
            CORE_ERROR("Failed to create export workbook: {}", output_filename);
            return false;
        }
        
        // 复制指定的工作表
        int exported_count = 0;
        for (const std::string& name : worksheet_names) {
            auto source_worksheet = getWorksheet(name);
            if (!source_worksheet) {
                CORE_WARN("Worksheet not found for export: {}", name);
                continue;
            }
            
            auto new_worksheet = export_workbook->addWorksheet(name);
            if (new_worksheet) {
                // 这里需要实现深拷贝逻辑
                // 简化版本：复制基本属性
                exported_count++;
                CORE_DEBUG("Exported worksheet: {}", name);
            }
        }
        
        // 复制文档属性
        export_workbook->doc_properties_ = doc_properties_;
        // 复制自定义属性
        for (const auto& prop : custom_property_manager_->getAllDetailedProperties()) {
            export_workbook->setCustomProperty(prop.name, prop.value);
        }
        
        // 保存导出的工作簿
        bool success = export_workbook->save();
        export_workbook->close();
        
        if (success) {
            CORE_INFO("Successfully exported {} worksheets to: {}", exported_count, output_filename);
        } else {
            CORE_ERROR("Failed to save exported workbook: {}", output_filename);
        }
        
        return success;
        
    } catch (const std::exception& e) {
        CORE_ERROR("Exception during worksheet export: {}", e.what());
        return false;
    }
}

int Workbook::batchRenameWorksheets(const std::unordered_map<std::string, std::string>& rename_map) {
    int renamed_count = 0;
    
    for (const auto& [old_name, new_name] : rename_map) {
        if (renameWorksheet(old_name, new_name)) {
            renamed_count++;
            CORE_DEBUG("Renamed worksheet: {} -> {}", old_name, new_name);
        } else {
            CORE_WARN("Failed to rename worksheet: {} -> {}", old_name, new_name);
        }
    }
    
    CORE_INFO("Batch rename completed: {} worksheets renamed", renamed_count);
    return renamed_count;
}

int Workbook::batchRemoveWorksheets(const std::vector<std::string>& worksheet_names) {
    int removed_count = 0;
    
    for (const std::string& name : worksheet_names) {
        if (removeWorksheet(name)) {
            removed_count++;
            CORE_DEBUG("Removed worksheet: {}", name);
        } else {
            CORE_WARN("Failed to remove worksheet: {}", name);
        }
    }
    
    CORE_INFO("Batch remove completed: {} worksheets removed", removed_count);
    return removed_count;
}

bool Workbook::reorderWorksheets(const std::vector<std::string>& new_order) {
    if (new_order.size() != worksheets_.size()) {
        CORE_ERROR("New order size ({}) doesn't match worksheet count ({})",
                 new_order.size(), worksheets_.size());
        return false;
    }
    
    try {
        std::vector<std::shared_ptr<Worksheet>> reordered_worksheets;
        reordered_worksheets.reserve(worksheets_.size());
        
        // 按新顺序重新排列工作表
        for (const std::string& name : new_order) {
            auto worksheet = getWorksheet(name);
            if (!worksheet) {
                CORE_ERROR("Worksheet not found in reorder list: {}", name);
                return false;
            }
            reordered_worksheets.push_back(worksheet);
        }
        
        // 替换工作表列表
        worksheets_ = std::move(reordered_worksheets);
        
        CORE_INFO("Successfully reordered {} worksheets", worksheets_.size());
        return true;
        
    } catch (const std::exception& e) {
        CORE_ERROR("Exception during worksheet reordering: {}", e.what());
        return false;
    }
}

int Workbook::findAndReplaceAll(const std::string& find_text, const std::string& replace_text,
                                const FindReplaceOptions& options) {
    int total_replacements = 0;
    
    for (const auto& worksheet : worksheets_) {
        // 检查工作表过滤器
        if (!options.worksheet_filter.empty()) {
            bool found = std::find(options.worksheet_filter.begin(), options.worksheet_filter.end(),
                                 worksheet->getName()) != options.worksheet_filter.end();
            if (!found) {
                continue; // 跳过不在过滤器中的工作表
            }
        }
        
        int replacements = worksheet->findAndReplace(find_text, replace_text,
                                                   options.match_case, options.match_entire_cell);
        total_replacements += replacements;
        
        if (replacements > 0) {
            CORE_DEBUG("Found and replaced {} occurrences in worksheet: {}",
                     replacements, worksheet->getName());
        }
    }
    
    CORE_INFO("Global find and replace completed: {} total replacements", total_replacements);
    return total_replacements;
}

std::vector<std::tuple<std::string, int, int>> Workbook::findAll(const std::string& search_text,
                                                                 const FindReplaceOptions& options) {
    std::vector<std::tuple<std::string, int, int>> results;
    
    for (const auto& worksheet : worksheets_) {
        // 检查工作表过滤器
        if (!options.worksheet_filter.empty()) {
            bool found = std::find(options.worksheet_filter.begin(), options.worksheet_filter.end(),
                                 worksheet->getName()) != options.worksheet_filter.end();
            if (!found) {
                continue; // 跳过不在过滤器中的工作表
            }
        }
        
        auto worksheet_results = worksheet->findCells(search_text, options.match_case, options.match_entire_cell);
        
        // 将结果添加到总结果中，包含工作表名称
        for (const auto& [row, col] : worksheet_results) {
            results.emplace_back(worksheet->getName(), row, col);
        }
        
        if (!worksheet_results.empty()) {
            CORE_DEBUG("Found {} matches in worksheet: {}", worksheet_results.size(), worksheet->getName());
        }
    }
    
    CORE_INFO("Global search completed: {} total matches found", results.size());
    return results;
}

Workbook::WorkbookStats Workbook::getStatistics() const {
    WorkbookStats stats;
    
    stats.total_worksheets = worksheets_.size();
    stats.total_formats = format_repo_->getFormatCount();
    
    // 计算总单元格数和内存使用
    for (const auto& worksheet : worksheets_) {
        size_t cell_count = worksheet->getCellCount();
        stats.total_cells += cell_count;
        stats.worksheet_cell_counts[worksheet->getName()] = cell_count;
        
        if (worksheet->isOptimizeMode()) {
            stats.memory_usage += worksheet->getMemoryUsage();
        }
    }
    
    // 估算工作簿本身的内存使用
    stats.memory_usage += sizeof(Workbook);
    stats.memory_usage += worksheets_.capacity() * sizeof(std::shared_ptr<Worksheet>);
    stats.memory_usage += format_repo_->getMemoryUsage();
    stats.memory_usage += custom_property_manager_->size() * sizeof(CustomProperty);
    stats.memory_usage += defined_name_manager_->size() * sizeof(DefinedName);
    
    return stats;
}

// ========== 智能模式选择辅助方法 ==========

size_t Workbook::estimateMemoryUsage() const {
    size_t total_memory = 0;
    
    // 估算工作表内存使用
    for (const auto& worksheet : worksheets_) {
        if (worksheet->isOptimizeMode()) {
            total_memory += worksheet->getMemoryUsage();
        } else {
            // 估算标准模式的内存使用
            auto [max_row, max_col] = worksheet->getUsedRange();
            if (max_row >= 0 && max_col >= 0) {
                size_t cell_count = (max_row + 1) * (max_col + 1);
                total_memory += cell_count * 100; // 估算每个单元格100字节
            }
        }
    }
    
    // 估算格式池内存
    total_memory += format_repo_->getMemoryUsage();
    
    // 估算共享字符串内存
    if (shared_string_table_) {
        total_memory += shared_string_table_->getMemoryUsage();
    }
    
    // 估算XML生成时的临时内存（约为数据的2-3倍）
    total_memory *= 3;
    
    return total_memory;
}

size_t Workbook::getTotalCellCount() const {
    size_t total_cells = 0;
    
    for (const auto& worksheet : worksheets_) {
        if (worksheet->isOptimizeMode()) {
            total_cells += worksheet->getCellCount();
        } else {
            auto [max_row, max_col] = worksheet->getUsedRange();
            if (max_row >= 0 && max_col >= 0) {
                // 估算实际有数据的单元格数量（不是整个矩形区域）
                size_t estimated_cells = 0;
                for (int row = 0; row <= max_row; ++row) {
                    for (int col = 0; col <= max_col; ++col) {
                        if (worksheet->hasCellAt(row, col)) {
                            estimated_cells++;
                        }
                    }
                }
                total_cells += estimated_cells;
            }
        }
    }
    
    return total_cells;
}

std::unique_ptr<StyleTransferContext> Workbook::copyStylesFrom(const Workbook& source_workbook) {
    CORE_DEBUG("开始从源工作簿复制样式数据");
    
    // 创建样式传输上下文
    auto transfer_context = std::make_unique<StyleTransferContext>(*source_workbook.format_repo_, *format_repo_);
    
    // 预加载所有映射以触发批量复制
    transfer_context->preloadAllMappings();
    
    auto stats = transfer_context->getTransferStats();
    CORE_DEBUG("完成样式复制，传输了{}个格式，去重了{}个", 
             stats.transferred_count, stats.deduplicated_count);
    
    // 🔧 关键修复：自动复制主题XML以保持颜色和字体一致性
    const std::string& source_theme = source_workbook.getThemeXML();
    if (!source_theme.empty()) {
        // 只有当前工作簿没有自定义主题时才复制源主题
        if (theme_xml_.empty()) {
            theme_xml_ = source_theme;
            CORE_DEBUG("自动复制主题XML ({} 字节)", theme_xml_.size());
        } else {
            CORE_DEBUG("当前工作簿已有自定义主题，保持现有主题不变");
        }
    } else {
        CORE_DEBUG("源工作簿无自定义主题，保持默认主题");
    }
    
    return transfer_context;
}

FormatRepository::DeduplicationStats Workbook::getStyleStats() const {
    return format_repo_->getDeduplicationStats();
}

bool Workbook::generateWithGenerator(bool use_streaming_writer) {
    if (!file_manager_) {
        CORE_ERROR("FileManager is null - cannot write workbook");
        return false;
    }
    std::unique_ptr<IFileWriter> writer;
    if (use_streaming_writer) {
        writer = std::make_unique<StreamingFileWriter>(file_manager_.get());
    } else {
        writer = std::make_unique<BatchFileWriter>(file_manager_.get());
    }
    ExcelStructureGenerator generator(this, std::move(writer));
    return generator.generate();
}

bool Workbook::isModified() const {
    // 检查DirtyManager是否有修改标记
    if (dirty_manager_ && dirty_manager_->hasDirtyData()) {
        return true;
    }
    
    // 检查主题是否被修改
    if (theme_dirty_) {
        return true;
    }
    
    // 检查工作表是否有修改（如果有hasChanges方法）
    for (const auto& worksheet : worksheets_) {
        if (worksheet) {
            // TODO: 检查Worksheet是否有hasChanges或类似方法
            // if (worksheet->hasChanges()) {
            //     return true;
            // }
        }
    }
    
    return false;
}

// ========== 访问模式检查辅助方法实现 ==========

void Workbook::ensureEditable(const std::string& operation) const {
    if (state_ == WorkbookState::READING) {
        std::string msg = "Cannot perform operation";
        if (!operation.empty()) {
            msg += " '" + operation + "'";
        }
        msg += ": workbook is opened in read-only mode. Use openForEditing() instead of openForReading().";
        
        CORE_ERROR("{}", msg);
        throw OperationException(msg, operation);
    }
}

void Workbook::ensureReadable(const std::string& operation) const {
    // 读取操作在任何模式下都是允许的
    // 这个方法预留用于未来可能的扩展，比如检查文件是否损坏等
    (void)operation; // 避免未使用参数警告
}

// ========== 🔧 新状态管理系统实现 ==========

bool Workbook::isStateValid(WorkbookState required_state) const {
    // 状态层级：CLOSED < CREATING/READING/EDITING
    // CREATING/READING/EDITING 是平级的，但有不同的权限
    
    switch (required_state) {
        case WorkbookState::CLOSED:
            return true; // 任何状态都可以关闭
            
        case WorkbookState::CREATING:
            return state_ == WorkbookState::CREATING;
            
        case WorkbookState::READING:
            // 读取操作在 READING 和 EDITING 状态都允许
            return state_ == WorkbookState::READING || state_ == WorkbookState::EDITING;
            
        case WorkbookState::EDITING:
            // 编辑操作只在 EDITING 和 CREATING 状态允许
            return state_ == WorkbookState::EDITING || state_ == WorkbookState::CREATING;
            
        default:
            return false;
    }
}

void Workbook::transitionToState(WorkbookState new_state, const std::string& reason) {
    if (state_ == new_state) {
        return; // 状态未改变
    }
    
    WorkbookState old_state = state_;
    state_ = new_state;
    
    CORE_DEBUG("Workbook state transition: {} -> {} ({})", 
              static_cast<int>(old_state), 
              static_cast<int>(new_state), 
              reason.empty() ? "no reason" : reason);
}

// ========== 样式构建器 ==========

StyleBuilder Workbook::createStyleBuilder() const {
    return StyleBuilder();
}

}} // namespace fastexcel::core
