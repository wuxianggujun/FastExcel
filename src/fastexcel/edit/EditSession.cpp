#include "fastexcel/edit/EditSession.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Workbook.hpp"
#include <algorithm>
#include <chrono>

namespace fastexcel::edit {

// 创建新的工作簿
std::unique_ptr<EditSession> EditSession::createNew(const core::Path& path) {
    auto session = std::make_unique<EditSession>();
    session->path_ = path;
    session->workbook_ = std::make_unique<core::Workbook>();
    session->dirty_manager_ = std::make_unique<core::DirtyManager>();
    session->is_new_file_ = true;
    session->auto_save_enabled_ = false;
    session->auto_save_interval_ = std::chrono::minutes(5);
    
    LOG_INFO("Created new EditSession for: {}", path.string());
    return session;
}

// 从现有文件创建
std::unique_ptr<EditSession> EditSession::fromExistingFile(const core::Path& path) {
    auto session = std::make_unique<EditSession>();
    session->path_ = path;
    session->workbook_ = std::make_unique<core::Workbook>();
    session->dirty_manager_ = std::make_unique<core::DirtyManager>();
    session->is_new_file_ = false;
    session->auto_save_enabled_ = false;
    session->auto_save_interval_ = std::chrono::minutes(5);
    
    // 打开文件
    auto result = session->workbook_->open(path);
    if (result != core::ErrorCode::Ok) {
        throw std::runtime_error("Failed to open file: " + path.string());
    }
    
    // 获取工作表列表
    session->worksheet_names_ = session->workbook_->getWorksheetNames();
    
    LOG_INFO("Opened EditSession for: {}, worksheets: {}", 
             path.string(), session->worksheet_names_.size());
    return session;
}

// 从只读工作簿开始编辑
std::unique_ptr<EditSession> EditSession::beginEdit(
    const read::ReadWorkbook& read_workbook,
    const core::Path& target_path) {
    
    auto session = std::make_unique<EditSession>();
    session->path_ = target_path;
    session->workbook_ = std::make_unique<core::Workbook>();
    session->dirty_manager_ = std::make_unique<core::DirtyManager>();
    session->is_new_file_ = (target_path != read_workbook.getPath());
    session->auto_save_enabled_ = false;
    session->auto_save_interval_ = std::chrono::minutes(5);
    
    // 复制只读工作簿的内容
    auto source_path = read_workbook.getPath();
    auto result = session->workbook_->open(source_path);
    if (result != core::ErrorCode::Ok) {
        throw std::runtime_error("Failed to copy from read-only workbook");
    }
    
    // 如果目标路径不同，标记为已修改
    if (session->is_new_file_) {
        session->dirty_manager_->markModified();
    }
    
    // 获取工作表列表
    session->worksheet_names_ = session->workbook_->getWorksheetNames();
    
    LOG_INFO("Started edit session from {} to {}", 
             source_path.string(), target_path.string());
    return session;
}

// 构造函数
EditSession::EditSession()
    : save_count_(0)
    , modification_count_(0) {
}

// 析构函数
EditSession::~EditSession() {
    // 如果有未保存的更改，发出警告
    if (hasUnsavedChanges()) {
        LOG_WARN("EditSession destroyed with unsaved changes: {}", path_.string());
    }
    
    // 停止自动保存
    stopAutoSave();
    
    LOG_DEBUG("EditSession closed: {}, saves: {}, modifications: {}", 
             path_.string(), save_count_, modification_count_);
}

// 获取工作表数量
size_t EditSession::getWorksheetCount() const {
    return worksheet_names_.size();
}

// 获取工作表名称列表
std::vector<std::string> EditSession::getWorksheetNames() const {
    return worksheet_names_;
}

// 获取只读工作表
std::unique_ptr<read::IReadOnlyWorksheet> EditSession::getWorksheet(size_t index) const {
    if (index >= worksheet_names_.size()) {
        LOG_ERROR("Worksheet index out of range: {}", index);
        return nullptr;
    }
    
    return getWorksheet(worksheet_names_[index]);
}

// 获取只读工作表（按名称）
std::unique_ptr<read::IReadOnlyWorksheet> EditSession::getWorksheet(const std::string& name) const {
    auto ws = workbook_->getWorksheet(name);
    if (!ws) {
        LOG_ERROR("Worksheet not found: {}", name);
        return nullptr;
    }
    
    // 包装为只读接口
    return std::make_unique<WorksheetReadWrapper>(ws);
}

// 检查工作表是否存在
bool EditSession::hasWorksheet(const std::string& name) const {
    return std::find(worksheet_names_.begin(), worksheet_names_.end(), name) 
           != worksheet_names_.end();
}

// 获取元数据
reader::WorkbookMetadata EditSession::getMetadata() const {
    reader::WorkbookMetadata metadata;
    metadata.worksheet_count = worksheet_names_.size();
    metadata.has_shared_strings = workbook_->hasSharedStrings();
    metadata.has_styles = workbook_->hasStyles();
    return metadata;
}

// 获取文件路径
core::Path EditSession::getPath() const {
    return path_;
}

// 获取访问模式
read::WorkbookAccessMode EditSession::getAccessMode() const {
    return read::WorkbookAccessMode::EDITABLE;
}

// 获取可编辑工作表
std::shared_ptr<IEditableWorksheet> EditSession::getWorksheetForEdit(const std::string& name) {
    auto ws = workbook_->getWorksheet(name);
    if (!ws) {
        LOG_ERROR("Worksheet not found: {}", name);
        return nullptr;
    }
    
    // 创建可编辑包装器
    auto editable = std::make_shared<EditableWorksheetImpl>(ws, dirty_manager_.get());
    
    // 记录修改
    modification_count_++;
    
    return editable;
}

// 添加新工作表
std::shared_ptr<IEditableWorksheet> EditSession::addWorksheet(const std::string& name) {
    // 检查名称是否已存在
    if (hasWorksheet(name)) {
        LOG_ERROR("Worksheet already exists: {}", name);
        return nullptr;
    }
    
    // 创建新工作表
    auto ws = workbook_->addWorksheet(name);
    if (!ws) {
        LOG_ERROR("Failed to create worksheet: {}", name);
        return nullptr;
    }
    
    // 更新工作表列表
    worksheet_names_.push_back(name);
    
    // 标记为已修改
    dirty_manager_->markModified();
    dirty_manager_->markWorksheetAdded(name);
    modification_count_++;
    
    LOG_INFO("Added worksheet: {}", name);
    
    // 返回可编辑包装器
    return std::make_shared<EditableWorksheetImpl>(ws, dirty_manager_.get());
}

// 删除工作表
bool EditSession::removeWorksheet(const std::string& name) {
    // 检查工作表是否存在
    if (!hasWorksheet(name)) {
        LOG_ERROR("Worksheet not found: {}", name);
        return false;
    }
    
    // 不能删除最后一个工作表
    if (worksheet_names_.size() == 1) {
        LOG_ERROR("Cannot remove the last worksheet");
        return false;
    }
    
    // 删除工作表
    if (!workbook_->removeWorksheet(name)) {
        LOG_ERROR("Failed to remove worksheet: {}", name);
        return false;
    }
    
    // 更新工作表列表
    worksheet_names_.erase(
        std::remove(worksheet_names_.begin(), worksheet_names_.end(), name),
        worksheet_names_.end()
    );
    
    // 标记为已修改
    dirty_manager_->markModified();
    dirty_manager_->markWorksheetRemoved(name);
    modification_count_++;
    
    LOG_INFO("Removed worksheet: {}", name);
    return true;
}

// 重命名工作表
bool EditSession::renameWorksheet(const std::string& old_name, const std::string& new_name) {
    // 检查旧名称是否存在
    if (!hasWorksheet(old_name)) {
        LOG_ERROR("Worksheet not found: {}", old_name);
        return false;
    }
    
    // 检查新名称是否已存在
    if (hasWorksheet(new_name)) {
        LOG_ERROR("Worksheet already exists: {}", new_name);
        return false;
    }
    
    // 重命名工作表
    if (!workbook_->renameWorksheet(old_name, new_name)) {
        LOG_ERROR("Failed to rename worksheet: {} -> {}", old_name, new_name);
        return false;
    }
    
    // 更新工作表列表
    auto it = std::find(worksheet_names_.begin(), worksheet_names_.end(), old_name);
    if (it != worksheet_names_.end()) {
        *it = new_name;
    }
    
    // 标记为已修改
    dirty_manager_->markModified();
    dirty_manager_->markWorksheetModified(new_name);
    modification_count_++;
    
    LOG_INFO("Renamed worksheet: {} -> {}", old_name, new_name);
    return true;
}

// 保存文件
bool EditSession::save() {
    return saveAs(path_);
}

// 另存为
bool EditSession::saveAs(const core::Path& path) {
    try {
        // 根据脏标记管理器决定保存策略
        auto strategy = dirty_manager_->getSaveStrategy();
        
        LOG_INFO("Saving with strategy: {}", static_cast<int>(strategy));
        
        bool success = false;
        switch (strategy) {
            case core::SaveStrategy::NONE:
                // 没有修改，不需要保存
                LOG_INFO("No changes to save");
                return true;
                
            case core::SaveStrategy::MINIMAL_UPDATE:
                // 最小更新（只更新修改的单元格）
                success = workbook_->saveIncremental(path);
                break;
                
            case core::SaveStrategy::SMART_EDIT:
                // 智能编辑（重新生成修改的工作表）
                success = workbook_->saveWithSmartEdit(path, 
                    dirty_manager_->getModifiedWorksheets());
                break;
                
            case core::SaveStrategy::FULL_REBUILD:
                // 完全重建
                success = workbook_->saveWithFullGeneration(path);
                break;
                
            default:
                LOG_ERROR("Unknown save strategy: {}", static_cast<int>(strategy));
                return false;
        }
        
        if (success) {
            // 清除脏标记
            dirty_manager_->clearAll();
            save_count_++;
            last_save_time_ = std::chrono::steady_clock::now();
            
            // 更新路径（如果是另存为）
            if (path != path_) {
                path_ = path;
                is_new_file_ = false;
            }
            
            LOG_INFO("File saved successfully: {}", path.string());
            return true;
        } else {
            LOG_ERROR("Failed to save file: {}", path.string());
            return false;
        }
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception during save: {}", e.what());
        return false;
    }
}

// 检查是否有未保存的更改
bool EditSession::hasUnsavedChanges() const {
    return dirty_manager_->isDirty();
}

// 获取修改的工作表列表
std::vector<std::string> EditSession::getModifiedWorksheets() const {
    return dirty_manager_->getModifiedWorksheets();
}

// 撤销所有更改
void EditSession::discardChanges() {
    if (!hasUnsavedChanges()) {
        return;
    }
    
    // 重新加载文件
    if (!is_new_file_) {
        workbook_->close();
        workbook_->open(path_);
        worksheet_names_ = workbook_->getWorksheetNames();
    } else {
        // 新文件，清空所有内容
        workbook_ = std::make_unique<core::Workbook>();
        worksheet_names_.clear();
    }
    
    // 清除脏标记
    dirty_manager_->clearAll();
    modification_count_ = 0;
    
    LOG_INFO("Discarded all changes");
}

// 启用自动保存
void EditSession::enableAutoSave(std::chrono::seconds interval) {
    auto_save_interval_ = interval;
    auto_save_enabled_ = true;
    
    // 启动自动保存线程
    startAutoSave();
    
    LOG_INFO("Auto-save enabled with interval: {} seconds", interval.count());
}

// 禁用自动保存
void EditSession::disableAutoSave() {
    auto_save_enabled_ = false;
    
    // 停止自动保存线程
    stopAutoSave();
    
    LOG_INFO("Auto-save disabled");
}

// 创建行写入器
std::unique_ptr<RowWriter> EditSession::createRowWriter(const std::string& worksheet_name) {
    // 获取或创建工作表
    auto ws = hasWorksheet(worksheet_name) 
              ? workbook_->getWorksheet(worksheet_name)
              : workbook_->addWorksheet(worksheet_name);
    
    if (!ws) {
        LOG_ERROR("Failed to get/create worksheet for RowWriter: {}", worksheet_name);
        return nullptr;
    }
    
    // 更新工作表列表
    if (!hasWorksheet(worksheet_name)) {
        worksheet_names_.push_back(worksheet_name);
    }
    
    // 创建行写入器
    return std::make_unique<RowWriter>(ws, dirty_manager_.get());
}

// 获取会话统计信息
EditSession::SessionStats EditSession::getStats() const {
    SessionStats stats;
    stats.save_count = save_count_;
    stats.modification_count = modification_count_;
    stats.worksheet_count = worksheet_names_.size();
    stats.has_unsaved_changes = hasUnsavedChanges();
    stats.auto_save_enabled = auto_save_enabled_;
    
    if (last_save_time_.time_since_epoch().count() > 0) {
        auto now = std::chrono::steady_clock::now();
        stats.seconds_since_last_save = 
            std::chrono::duration_cast<std::chrono::seconds>(now - last_save_time_).count();
    } else {
        stats.seconds_since_last_save = -1;
    }
    
    return stats;
}

// 启动自动保存
void EditSession::startAutoSave() {
    if (!auto_save_enabled_ || auto_save_thread_.joinable()) {
        return;
    }
    
    auto_save_running_ = true;
    auto_save_thread_ = std::thread([this]() {
        while (auto_save_running_) {
            std::this_thread::sleep_for(auto_save_interval_);
            
            if (auto_save_running_ && hasUnsavedChanges()) {
                LOG_DEBUG("Auto-saving...");
                if (save()) {
                    LOG_INFO("Auto-save completed");
                } else {
                    LOG_ERROR("Auto-save failed");
                }
            }
        }
    });
}

// 停止自动保存
void EditSession::stopAutoSave() {
    if (!auto_save_thread_.joinable()) {
        return;
    }
    
    auto_save_running_ = false;
    auto_save_thread_.join();
}

// EditableWorksheetImpl 实现
EditableWorksheetImpl::EditableWorksheetImpl(
    std::shared_ptr<core::Worksheet> worksheet,
    core::DirtyManager* dirty_manager)
    : worksheet_(worksheet)
    , dirty_manager_(dirty_manager) {
}

void EditableWorksheetImpl::writeString(size_t row, size_t col, const std::string& value) {
    worksheet_->setCellString(row, col, value);
    dirty_manager_->markCellModified(worksheet_->getName(), row, col);
}

void EditableWorksheetImpl::writeNumber(size_t row, size_t col, double value) {
    worksheet_->setCellNumber(row, col, value);
    dirty_manager_->markCellModified(worksheet_->getName(), row, col);
}

void EditableWorksheetImpl::writeBool(size_t row, size_t col, bool value) {
    worksheet_->setCellBool(row, col, value);
    dirty_manager_->markCellModified(worksheet_->getName(), row, col);
}

void EditableWorksheetImpl::writeFormula(size_t row, size_t col, const std::string& formula) {
    worksheet_->setCellFormula(row, col, formula);
    dirty_manager_->markCellModified(worksheet_->getName(), row, col);
}

void EditableWorksheetImpl::clearCell(size_t row, size_t col) {
    worksheet_->clearCell(row, col);
    dirty_manager_->markCellModified(worksheet_->getName(), row, col);
}

// WorksheetReadWrapper 实现
WorksheetReadWrapper::WorksheetReadWrapper(std::shared_ptr<core::Worksheet> worksheet)
    : worksheet_(worksheet) {
}

std::string WorksheetReadWrapper::getName() const {
    return worksheet_->getName();
}

size_t WorksheetReadWrapper::getRowCount() const {
    return worksheet_->getRowCount();
}

size_t WorksheetReadWrapper::getColumnCount() const {
    return worksheet_->getColumnCount();
}

std::string WorksheetReadWrapper::readString(size_t row, size_t col) const {
    auto cell = worksheet_->getCell(row, col);
    return cell ? cell->getString() : "";
}

double WorksheetReadWrapper::readNumber(size_t row, size_t col) const {
    auto cell = worksheet_->getCell(row, col);
    return cell ? cell->getNumber() : 0.0;
}

bool WorksheetReadWrapper::readBool(size_t row, size_t col) const {
    auto cell = worksheet_->getCell(row, col);
    return cell ? cell->getBool() : false;
}

} // namespace fastexcel::edit
