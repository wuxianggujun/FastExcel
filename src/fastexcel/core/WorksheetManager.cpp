#include "fastexcel/core/WorksheetManager.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>

namespace fastexcel {
namespace core {

WorksheetManager::WorksheetManager(Workbook* workbook)
    : workbook_(workbook) {
    if (!workbook_) {
        throw std::invalid_argument("Workbook cannot be null");
    }
}

WorksheetManager::~WorksheetManager() {
    clear();
}

// 创建和添加

WorksheetManager::WorksheetPtr WorksheetManager::createWorksheet(const std::string& name) {
    std::string worksheet_name = name.empty() ? generateUniqueName() : name;
    
    if (!validateName(worksheet_name)) {
        FASTEXCEL_LOG_ERROR("Invalid worksheet name: " + worksheet_name);
        return nullptr;
    }
    
    if (exists(worksheet_name)) {
        FASTEXCEL_LOG_ERROR("Worksheet already exists: " + worksheet_name);
        return nullptr;
    }
    
    if (worksheets_.size() >= config_.max_sheets) {
        FASTEXCEL_LOG_ERROR("Maximum number of sheets reached: " + std::to_string(config_.max_sheets));
        return nullptr;
    }
    
    // 创建新工作表
    auto worksheet = std::make_shared<Worksheet>(
        worksheet_name,
        std::shared_ptr<Workbook>(workbook_, [](Workbook*){}),
        next_sheet_id_++
    );
    
    // 添加到集合
    worksheets_.push_back(worksheet);
    
    // 更新索引
    size_t index = worksheets_.size() - 1;
    name_index_[worksheet_name] = index;
    id_index_[worksheet->getSheetId()] = index;
    
    stats_.total_created++;
    FASTEXCEL_LOG_DEBUG("Created worksheet: {} (ID: {})", worksheet_name, worksheet->getSheetId());
    
    return worksheet;
}

bool WorksheetManager::addWorksheet(WorksheetPtr worksheet) {
    if (!worksheet) {
        FASTEXCEL_LOG_ERROR("Cannot add null worksheet");
        return false;
    }
    
    const std::string& name = worksheet->getName();
    
    if (!validateName(name)) {
        FASTEXCEL_LOG_ERROR("Invalid worksheet name: " + name);
        return false;
    }
    
    if (exists(name)) {
        FASTEXCEL_LOG_ERROR("Worksheet already exists: " + name);
        return false;
    }
    
    if (worksheets_.size() >= config_.max_sheets) {
        FASTEXCEL_LOG_ERROR("Maximum number of sheets reached");
        return false;
    }
    
    worksheets_.push_back(worksheet);
    
    size_t index = worksheets_.size() - 1;
    name_index_[name] = index;
    id_index_[worksheet->getSheetId()] = index;
    
    stats_.total_created++;
    return true;
}

std::vector<WorksheetManager::WorksheetPtr> WorksheetManager::createWorksheets(
    size_t count, const std::string& name_prefix) {
    
    std::vector<WorksheetPtr> created;
    created.reserve(count);
    
    std::string prefix = name_prefix.empty() ? config_.default_name_prefix : name_prefix;
    
    for (size_t i = 0; i < count; ++i) {
        auto worksheet = createWorksheet(generateUniqueName(prefix));
        if (worksheet) {
            created.push_back(worksheet);
        }
    }
    
    return created;
}

// 查找和访问

WorksheetManager::WorksheetPtr WorksheetManager::getByName(const std::string& name) {
    stats_.lookups_by_name++;
    
    auto it = name_index_.find(name);
    if (it != name_index_.end() && isValidIndex(it->second)) {
        return worksheets_[it->second];
    }
    return nullptr;
}

WorksheetManager::ConstWorksheetPtr WorksheetManager::getByName(const std::string& name) const {
    stats_.lookups_by_name++;
    
    auto it = name_index_.find(name);
    if (it != name_index_.end() && isValidIndex(it->second)) {
        return worksheets_[it->second];
    }
    return nullptr;
}

WorksheetManager::WorksheetPtr WorksheetManager::getById(int id) {
    stats_.lookups_by_id++;
    
    auto it = id_index_.find(id);
    if (it != id_index_.end() && isValidIndex(it->second)) {
        return worksheets_[it->second];
    }
    return nullptr;
}

WorksheetManager::ConstWorksheetPtr WorksheetManager::getById(int id) const {
    stats_.lookups_by_id++;
    
    auto it = id_index_.find(id);
    if (it != id_index_.end() && isValidIndex(it->second)) {
        return worksheets_[it->second];
    }
    return nullptr;
}

WorksheetManager::WorksheetPtr WorksheetManager::getByIndex(size_t index) {
    stats_.lookups_by_index++;
    
    if (isValidIndex(index)) {
        return worksheets_[index];
    }
    return nullptr;
}

WorksheetManager::ConstWorksheetPtr WorksheetManager::getByIndex(size_t index) const {
    stats_.lookups_by_index++;
    
    if (isValidIndex(index)) {
        return worksheets_[index];
    }
    return nullptr;
}

std::vector<WorksheetManager::ConstWorksheetPtr> WorksheetManager::getAll() const {
    std::vector<ConstWorksheetPtr> result;
    result.reserve(worksheets_.size());
    
    for (const auto& ws : worksheets_) {
        result.push_back(ws);
    }
    
    return result;
}

std::vector<WorksheetManager::WorksheetPtr> WorksheetManager::findWhere(WorksheetPredicate predicate) {
    std::vector<WorksheetPtr> result;
    
    for (const auto& ws : worksheets_) {
        if (predicate(ws)) {
            result.push_back(ws);
        }
    }
    
    return result;
}

// 删除操作

bool WorksheetManager::removeByName(const std::string& name) {
    auto it = name_index_.find(name);
    if (it == name_index_.end()) {
        return false;
    }
    
    return removeByIndex(it->second);
}

bool WorksheetManager::removeById(int id) {
    auto it = id_index_.find(id);
    if (it == id_index_.end()) {
        return false;
    }
    
    return removeByIndex(it->second);
}

bool WorksheetManager::removeByIndex(size_t index) {
    if (!isValidIndex(index)) {
        return false;
    }
    
    auto worksheet = worksheets_[index];
    
    // 从索引中移除
    removeFromIndexes(worksheet);
    
    // 从列表中移除
    worksheets_.erase(worksheets_.begin() + index);
    
    // 更新后续索引
    updateIndexes(index);
    
    // 调整活动索引
    if (active_index_ >= worksheets_.size() && !worksheets_.empty()) {
        active_index_ = worksheets_.size() - 1;
    }
    
    stats_.total_deleted++;
    FASTEXCEL_LOG_DEBUG("Removed worksheet: {} (ID: {})", worksheet->getName(), worksheet->getSheetId());
    
    return true;
}

size_t WorksheetManager::clear() {
    size_t count = worksheets_.size();
    
    worksheets_.clear();
    name_index_.clear();
    id_index_.clear();
    next_sheet_id_ = 1;
    active_index_ = 0;
    
    stats_.total_deleted += count;
    
    return count;
}

// 重命名和移动

bool WorksheetManager::rename(const std::string& old_name, const std::string& new_name) {
    if (!validateName(new_name)) {
        return false;
    }
    
    if (exists(new_name)) {
        return false;
    }
    
    auto it = name_index_.find(old_name);
    if (it == name_index_.end()) {
        return false;
    }
    
    size_t index = it->second;
    auto worksheet = worksheets_[index];
    
    // 更新工作表名称
    worksheet->setName(new_name);
    
    // 更新名称索引
    name_index_.erase(it);
    name_index_[new_name] = index;
    
    FASTEXCEL_LOG_DEBUG("Renamed worksheet: {} -> {}", old_name, new_name);
    return true;
}

bool WorksheetManager::move(size_t from_index, size_t to_index) {
    if (!isValidIndex(from_index) || !isValidIndex(to_index)) {
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
    
    // 重建索引
    rebuildIndexes();
    
    FASTEXCEL_LOG_DEBUG("Moved worksheet from index {} to {}", from_index, to_index);
    return true;
}

bool WorksheetManager::swap(size_t index1, size_t index2) {
    if (!isValidIndex(index1) || !isValidIndex(index2)) {
        return false;
    }
    
    if (index1 == index2) {
        return true;
    }
    
    std::swap(worksheets_[index1], worksheets_[index2]);
    
    // 重建索引
    rebuildIndexes();
    
    return true;
}

// 复制和克隆

WorksheetManager::WorksheetPtr WorksheetManager::copy(
    const std::string& source_name, const std::string& new_name) {
    
    auto source = getByName(source_name);
    if (!source) {
        return nullptr;
    }
    
    return deepClone(source, new_name);
}

WorksheetManager::WorksheetPtr WorksheetManager::deepClone(
    const WorksheetPtr& source, const std::string& new_name) {
    
    if (!source) {
        return nullptr;
    }
    
    if (!validateName(new_name) || exists(new_name)) {
        return nullptr;
    }
    
    // 创建新工作表
    auto cloned = createWorksheet(new_name);
    if (!cloned) {
        return nullptr;
    }
    
    // TODO: 实现深度复制逻辑
    // 这里应该复制所有单元格、样式、设置等
    
    return cloned;
}

// 活动工作表管理

void WorksheetManager::setActive(size_t index) {
    if (!isValidIndex(index)) {
        return;
    }
    
    // 取消所有工作表的选中状态
    for (auto& ws : worksheets_) {
        ws->setTabSelected(false);
    }
    
    // 设置新的活动工作表
    worksheets_[index]->setTabSelected(true);
    active_index_ = index;
}

bool WorksheetManager::setActive(const std::string& name) {
    auto it = name_index_.find(name);
    if (it == name_index_.end()) {
        return false;
    }
    
    setActive(it->second);
    return true;
}

WorksheetManager::WorksheetPtr WorksheetManager::getActive() {
    if (worksheets_.empty()) {
        return nullptr;
    }
    
    if (active_index_ >= worksheets_.size()) {
        active_index_ = 0;
    }
    
    return worksheets_[active_index_];
}

WorksheetManager::ConstWorksheetPtr WorksheetManager::getActive() const {
    if (worksheets_.empty()) {
        return nullptr;
    }
    
    size_t safe_index = (active_index_ < worksheets_.size()) ? active_index_ : 0;
    return worksheets_[safe_index];
}

// 查询和统计

WorksheetManager::WorksheetInfo WorksheetManager::getInfo(size_t index) const {
    WorksheetInfo info{};
    
    if (isValidIndex(index)) {
        auto ws = worksheets_[index];
        info.id = ws->getSheetId();
        info.name = ws->getName();
        info.index = index;
        info.visible = true; // 假设所有工作表都可见，如需隐藏功能需要在Worksheet类中添加
        info.selected = ws->isTabSelected();
        info.row_count = ws->getRowCount();
        info.col_count = ws->getColumnCount();
    }
    
    return info;
}

std::vector<WorksheetManager::WorksheetInfo> WorksheetManager::getAllInfo() const {
    std::vector<WorksheetInfo> result;
    result.reserve(worksheets_.size());
    
    for (size_t i = 0; i < worksheets_.size(); ++i) {
        result.push_back(getInfo(i));
    }
    
    return result;
}

// 验证和工具

bool WorksheetManager::validateName(const std::string& name) const {
    if (name.empty()) {
        return false;
    }
    
    if (name.length() > 31) {  // Excel限制
        return false;
    }
    
    // 检查非法字符
    const std::string invalid_chars = ":\\/\\?*[]";
    if (name.find_first_of(invalid_chars) != std::string::npos) {
        return false;
    }
    
    return true;
}

std::string WorksheetManager::generateUniqueName(const std::string& prefix) {
    std::string base = prefix.empty() ? config_.default_name_prefix : prefix;
    std::string name;
    int counter = 1;
    
    do {
        name = base + std::to_string(counter++);
    } while (exists(name));
    
    return name;
}

void WorksheetManager::rebuildIndexes() {
    name_index_.clear();
    id_index_.clear();
    
    for (size_t i = 0; i < worksheets_.size(); ++i) {
        auto ws = worksheets_[i];
        name_index_[ws->getName()] = i;
        id_index_[ws->getSheetId()] = i;
    }
}

// 私有辅助方法

void WorksheetManager::updateIndexes(size_t start_index) {
    for (size_t i = start_index; i < worksheets_.size(); ++i) {
        auto ws = worksheets_[i];
        name_index_[ws->getName()] = i;
        id_index_[ws->getSheetId()] = i;
    }
}

void WorksheetManager::removeFromIndexes(const WorksheetPtr& worksheet) {
    if (worksheet) {
        name_index_.erase(worksheet->getName());
        id_index_.erase(worksheet->getSheetId());
    }
}

size_t WorksheetManager::getNextAvailableId() {
    return next_sheet_id_++;
}

}} // namespace fastexcel::core
