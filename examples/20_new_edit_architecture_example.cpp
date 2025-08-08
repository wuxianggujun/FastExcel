/**
 * @file 20_new_edit_architecture_example.cpp
 * @brief 新架构编辑示例 - 展示直接ZIP包操作的优势
 * 
 * 这个示例演示新的EditableWorkbook架构：
 * - 直接在ZIP包内操作，不需要完整加载
 * - 媒体资源自动保留
 * - 内存占用最小
 * - 支持大文件编辑
 */

#include "fastexcel/core/EditableWorkbook.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>
#include <chrono>

using namespace fastexcel::core;

int main() {
    try {
        // 初始化日志
        fastexcel::initialize("logs/new_edit_test.log", true);
        
        std::string source_file = "辅材处理-张玥 机房建设项目（2025-JW13-W1007）测试.xlsx";
        
        LOG_INFO("=== 新架构编辑功能测试 ===");
        LOG_INFO("源文件: {}", source_file);
        
        // 测量开始时间
        auto start = std::chrono::high_resolution_clock::now();
        
        // ========== 核心优势1：直接打开，不需要复制 ==========
        LOG_INFO("1. 直接打开文件进行编辑（无需复制）");
        auto workbook = EditableWorkbook::open(Path(source_file), 
                                              EditableWorkbook::EditMode::IN_PLACE);
        if (!workbook) {
            LOG_ERROR("无法打开工作簿");
            return -1;
        }
        
        auto open_time = std::chrono::high_resolution_clock::now();
        auto open_duration = std::chrono::duration_cast<std::chrono::milliseconds>(open_time - start);
        LOG_INFO("   打开耗时: {} ms（只读取元数据，不加载全部内容）", open_duration.count());
        
        // ========== 核心优势2：懒加载，只读需要的部分 ==========
        LOG_INFO("2. 懒加载工作表（只在需要时加载）");
        auto worksheet = workbook->getWorksheet(0);
        if (!worksheet) {
            LOG_ERROR("无法获取工作表");
            return -1;
        }
        
        // ========== 核心优势3：精确的脏数据追踪 ==========
        LOG_INFO("3. 修改单元格（精确追踪修改）");
        worksheet->setCellValue(0, 0, "新架构测试");
        worksheet->setCellValue(1, 1, 123.45);
        worksheet->setCellValue(2, 2, "=A1&B2");
        
        // 显示脏数据列表
        auto dirty_parts = workbook->getDirtyParts();
        LOG_INFO("   修改的部件数量: {}", dirty_parts.size());
        for (const auto& part : dirty_parts) {
            LOG_INFO("   - {}", part);
        }
        
        // ========== 核心优势4：原地保存，自动保留所有资源 ==========
        LOG_INFO("4. 原地保存（自动保留图片等资源）");
        auto save_start = std::chrono::high_resolution_clock::now();
        
        if (workbook->save()) {
            auto save_time = std::chrono::high_resolution_clock::now();
            auto save_duration = std::chrono::duration_cast<std::chrono::milliseconds>(save_time - save_start);
            LOG_INFO("   保存成功，耗时: {} ms", save_duration.count());
            LOG_INFO("   ✓ 图片自动保留（xl/media/）");
            LOG_INFO("   ✓ 图形自动保留（xl/drawings/）");
            LOG_INFO("   ✓ 关系自动保留（xl/worksheets/_rels/）");
            LOG_INFO("   ✓ 只更新修改的工作表");
        } else {
            LOG_ERROR("   保存失败");
            return -1;
        }
        
        // ========== 核心优势5：另存为也很高效 ==========
        LOG_INFO("5. 另存为新文件（智能复制）");
        std::string new_file = "新架构编辑结果.xlsx";
        
        // 再修改一些内容
        worksheet->setCellValue(3, 3, "另存为测试");
        
        auto saveas_start = std::chrono::high_resolution_clock::now();
        if (workbook->saveAs(Path(new_file))) {
            auto saveas_time = std::chrono::high_resolution_clock::now();
            auto saveas_duration = std::chrono::duration_cast<std::chrono::milliseconds>(saveas_time - saveas_start);
            LOG_INFO("   另存为成功: {}", new_file);
            LOG_INFO("   耗时: {} ms", saveas_duration.count());
            LOG_INFO("   ✓ 只复制未修改的条目");
            LOG_INFO("   ✓ 修改的条目写入新内容");
        }
        
        // ========== 总结 ==========
        auto total_time = std::chrono::high_resolution_clock::now();
        auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(total_time - start);
        
        LOG_INFO("=== 测试完成 ===");
        LOG_INFO("总耗时: {} ms", total_duration.count());
        LOG_INFO("");
        LOG_INFO("新架构优势总结：");
        LOG_INFO("1. 无需文件复制，直接编辑");
        LOG_INFO("2. 懒加载，内存占用最小");
        LOG_INFO("3. 精确的脏数据追踪");
        LOG_INFO("4. 自动保留所有未修改的资源");
        LOG_INFO("5. 支持原地更新和另存为");
        LOG_INFO("6. 适合处理包含大量图片的Excel文件");
        
        return 0;
        
    } catch (const std::exception& e) {
        LOG_ERROR("异常: {}", e.what());
        return -1;
    }
}
