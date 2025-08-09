#include "fastexcel/edit/EditSession.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Workbook.hpp"
#include <algorithm>
#include <chrono>

namespace fastexcel::edit {

// 创建新的工作簿
std::unique_ptr<EditSession> EditSession::createNew(const core::Path& path) {
    auto session = std::unique_ptr<EditSession>(new EditSession(path, read::WorkbookAccessMode::CREATE_NEW));
    session->initializeNewWorkbook();
    LOG_INFO("Created new EditSession for: {}", path.string());
    return session;
}

// 从现有文件创建
std::unique_ptr<EditSession> EditSession::fromExistingFile(const core::Path& path) {
    auto session = std::unique_ptr<EditSession>(new EditSession(path, read::WorkbookAccessMode::EDITABLE));
    if (!session->loadFromFile(path)) {
        throw std::runtime_error("Failed to open file: " + path.string());
    }
    LOG_INFO("Opened EditSession for: {}", path.string());
    return session;
}

// 从只读工作簿开始编辑
std::unique_ptr<EditSession> EditSession::beginEdit(
    const read::ReadWorkbook& read_workbook,
    const core::Path& target_path) {
    
    auto session = std::unique_ptr<EditSession>(new EditSession(target_path, read::WorkbookAccessMode::EDITABLE));
    session->initializeNewWorkbook();
    
    // TODO: 实现从ReadWorkbook复制数据的逻辑
    
    LOG_INFO("Started edit session from ReadWorkbook to: {}", target_path.string());
    return session;
}

// 私有构造函数
EditSession::EditSession(const core::Path& path, read::WorkbookAccessMode mode)
    : access_mode_(mode), filename_(path.string()) {
    
    workbook_ = core::Workbook::create(path);
    dirty_manager_ = std::make_unique<core::DirtyManager>();
    
    if (!workbook_->open()) {
        LOG_ERROR("Failed to open workbook: {}", path.string());
    }
}

// 析构函数
EditSession::~EditSession() {
    if (workbook_) {
        workbook_->close();
    }
    LOG_DEBUG("EditSession destroyed for: {}", filename_);
}

// ========== IReadOnlyWorkbook 接口实现 ==========

size_t EditSession::getWorksheetCount() const {
    return workbook_ ? workbook_->getWorksheetCount() : 0;
}

std::vector<std::string> EditSession::getWorksheetNames() const {
    return workbook_ ? workbook_->getWorksheetNames() : std::vector<std::string>();
}

std::shared_ptr<const read::ReadWorksheet> EditSession::getWorksheet(const std::string& name) const {
    // TODO: 实现ReadWorksheet包装
    return nullptr;
}

std::shared_ptr<const read::ReadWorksheet> EditSession::getWorksheet(size_t index) const {
    // TODO: 实现ReadWorksheet包装
    return nullptr;
}

bool EditSession::hasWorksheet(const std::string& name) const {
    return workbook_ && workbook_->getWorksheet(name) != nullptr;
}

int EditSession::getWorksheetIndex(const std::string& name) const {
    if (!workbook_) return -1;
    
    auto names = workbook_->getWorksheetNames();
    auto it = std::find(names.begin(), names.end(), name);
    return it != names.end() ? static_cast<int>(std::distance(names.begin(), it)) : -1;
}

// ========== IEditableWorkbook 接口实现 ==========

std::shared_ptr<EditWorksheet> EditSession::getWorksheetForEdit(const std::string& name) {
    // TODO: 实现EditWorksheet包装
    return nullptr;
}

std::shared_ptr<EditWorksheet> EditSession::getWorksheetForEdit(size_t index) {
    // TODO: 实现EditWorksheet包装
    return nullptr;
}

std::shared_ptr<EditWorksheet> EditSession::addWorksheet(const std::string& name) {
    if (!workbook_) return nullptr;
    
    auto worksheet = workbook_->addWorksheet(name);
    if (worksheet) {
        has_unsaved_changes_ = true;
        // TODO: 创建EditWorksheet包装
    }
    
    return nullptr;
}

bool EditSession::removeWorksheet(const std::string& name) {
    if (!workbook_) return false;
    
    bool success = workbook_->removeWorksheet(name);
    if (success) {
        has_unsaved_changes_ = true;
    }
    
    return success;
}

bool EditSession::removeWorksheet(size_t index) {
    if (!workbook_) return false;
    
    bool success = workbook_->removeWorksheet(index);
    if (success) {
        has_unsaved_changes_ = true;
    }
    
    return success;
}

bool EditSession::save() {
    if (!workbook_) return false;
    
    bool success = workbook_->save();
    if (success) {
        has_unsaved_changes_ = false;
    }
    
    return success;
}

bool EditSession::saveAs(const std::string& filename) {
    if (!workbook_) return false;
    
    bool success = workbook_->saveAs(filename);
    if (success) {
        filename_ = filename;
        has_unsaved_changes_ = false;
    }
    
    return success;
}

bool EditSession::hasUnsavedChanges() const {
    return has_unsaved_changes_;
}

void EditSession::markAsModified() {
    has_unsaved_changes_ = true;
}

void EditSession::discardChanges() {
    has_unsaved_changes_ = false;
    // TODO: 实现变更丢弃逻辑
}

bool EditSession::renameWorksheet(const std::string& old_name, const std::string& new_name) {
    if (!workbook_) return false;
    
    bool success = workbook_->renameWorksheet(old_name, new_name);
    if (success) {
        has_unsaved_changes_ = true;
    }
    
    return success;
}

bool EditSession::moveWorksheet(size_t from_index, size_t to_index) {
    if (!workbook_) return false;
    
    bool success = workbook_->moveWorksheet(from_index, to_index);
    if (success) {
        has_unsaved_changes_ = true;
    }
    
    return success;
}

// ========== 高性能编辑方法 ==========

std::unique_ptr<RowWriter> EditSession::createRowWriter(const std::string& worksheet_name) {
    // TODO: 实现RowWriter创建
    return nullptr;
}

void EditSession::setUseSharedStrings(bool enable) {
    if (workbook_) {
        workbook_->setUseSharedStrings(enable);
    }
}

void EditSession::setCompressionLevel(int level) {
    if (workbook_) {
        workbook_->setCompressionLevel(level);
    }
}

// ========== 私有方法 ==========

void EditSession::initializeNewWorkbook() {
    // 初始化新工作簿的基本设置
    has_unsaved_changes_ = true;
}

bool EditSession::loadFromFile(const core::Path& path) {
    if (!workbook_) {
        return false;
    }
    
    // 工作簿已经在构造函数中打开
    return true;
}

} // namespace fastexcel::edit