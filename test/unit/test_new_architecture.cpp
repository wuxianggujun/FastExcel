#include <gtest/gtest.h>
#include "fastexcel/FastExcel.hpp"
#include "fastexcel/read/ReadWorkbook.hpp"
#include "fastexcel/read/ReadWorksheet.hpp"
#include "fastexcel/edit/EditSession.hpp"
#include "fastexcel/edit/RowWriter.hpp"
#include <filesystem>
#include <chrono>
#include <random>

using namespace fastexcel;
namespace fs = std::filesystem;

class NewArchitectureTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建测试目录
        test_dir_ = fs::temp_directory_path() / "fastexcel_test";
        fs::create_directories(test_dir_);
        
        // 生成测试文件路径
        test_file_ = test_dir_ / "test_workbook.xlsx";
        readonly_file_ = test_dir_ / "readonly_test.xlsx";
        large_file_ = test_dir_ / "large_test.xlsx";
    }
    
    void TearDown() override {
        // 清理测试文件
        fs::remove_all(test_dir_);
    }
    
    fs::path test_dir_;
    fs::path test_file_;
    fs::path readonly_file_;
    fs::path large_file_;
};

// 测试1：创建新工作簿
TEST_F(NewArchitectureTest, CreateNewWorkbook) {
    // 使用工厂创建新工作簿
    auto workbook = FastExcel::createWorkbook(core::Path(test_file_.string()));
    ASSERT_NE(workbook, nullptr);
    
    // 添加工作表
    auto ws1 = workbook->addWorksheet("Sheet1");
    ASSERT_NE(ws1, nullptr);
    
    auto ws2 = workbook->addWorksheet("Data");
    ASSERT_NE(ws2, nullptr);
    
    // 写入数据
    ws1->writeString(0, 0, "Hello");
    ws1->writeNumber(0, 1, 42.5);
    ws1->writeBool(0, 2, true);
    ws1->writeFormula(0, 3, "=B1*2");
    
    ws2->writeString(0, 0, "Name");
    ws2->writeString(0, 1, "Value");
    ws2->writeString(1, 0, "Test");
    ws2->writeNumber(1, 1, 100);
    
    // 保存文件
    ASSERT_TRUE(workbook->save());
    
    // 验证文件存在
    ASSERT_TRUE(fs::exists(test_file_));
    
    // 验证工作表数量
    EXPECT_EQ(workbook->getWorksheetCount(), 2);
    
    // 验证工作表名称
    auto names = workbook->getWorksheetNames();
    EXPECT_EQ(names.size(), 2);
    EXPECT_EQ(names[0], "Sheet1");
    EXPECT_EQ(names[1], "Data");
}

// 测试2：只读模式访问
TEST_F(NewArchitectureTest, ReadOnlyAccess) {
    // 先创建一个文件
    {
        auto workbook = FastExcel::createWorkbook(core::Path(readonly_file_.string()));
        auto ws = workbook->addWorksheet("TestSheet");
        
        // 写入测试数据
        ws->writeString(0, 0, "Header1");
        ws->writeString(0, 1, "Header2");
        ws->writeString(0, 2, "Header3");
        
        for (int i = 1; i <= 10; ++i) {
            ws->writeString(i, 0, "Row" + std::to_string(i));
            ws->writeNumber(i, 1, i * 10.5);
            ws->writeBool(i, 2, i % 2 == 0);
        }
        
        ASSERT_TRUE(workbook->save());
    }
    
    // 以只读模式打开
    auto readonly = FastExcel::openForReading(core::Path(readonly_file_.string()));
    ASSERT_NE(readonly, nullptr);
    
    // 验证访问模式
    EXPECT_EQ(readonly->getAccessMode(), read::WorkbookAccessMode::READ_ONLY);
    
    // 获取工作表
    auto ws = readonly->getWorksheet("TestSheet");
    ASSERT_NE(ws, nullptr);
    
    // 读取数据
    EXPECT_EQ(ws->readString(0, 0), "Header1");
    EXPECT_EQ(ws->readString(0, 1), "Header2");
    EXPECT_EQ(ws->readString(0, 2), "Header3");
    
    EXPECT_EQ(ws->readString(1, 0), "Row1");
    EXPECT_DOUBLE_EQ(ws->readNumber(1, 1), 10.5);
    EXPECT_FALSE(ws->readBool(1, 2));
    
    EXPECT_EQ(ws->readString(2, 0), "Row2");
    EXPECT_DOUBLE_EQ(ws->readNumber(2, 1), 21.0);
    EXPECT_TRUE(ws->readBool(2, 2));
    
    // 验证行列数
    EXPECT_EQ(ws->getRowCount(), 11);
    EXPECT_EQ(ws->getColumnCount(), 3);
    
    // 测试行迭代器
    auto iterator = ws->createRowIterator();
    ASSERT_NE(iterator, nullptr);
    
    int row_count = 0;
    while (iterator->hasNext()) {
        auto row = iterator->next();
        row_count++;
        EXPECT_EQ(row.size(), 3);
    }
    EXPECT_EQ(row_count, 11);
}

// 测试3：编辑现有文件
TEST_F(NewArchitectureTest, EditExistingFile) {
    // 先创建一个文件
    {
        auto workbook = FastExcel::createWorkbook(core::Path(test_file_.string()));
        auto ws = workbook->addWorksheet("Original");
        ws->writeString(0, 0, "Original Data");
        ws->writeNumber(1, 0, 100);
        ASSERT_TRUE(workbook->save());
    }
    
    // 以编辑模式打开
    auto editable = FastExcel::openForEditing(core::Path(test_file_.string()));
    ASSERT_NE(editable, nullptr);
    
    // 获取现有工作表
    auto ws = editable->getWorksheetForEdit("Original");
    ASSERT_NE(ws, nullptr);
    
    // 修改数据
    ws->writeString(0, 0, "Modified Data");
    ws->writeNumber(1, 0, 200);
    ws->writeString(2, 0, "New Row");
    
    // 添加新工作表
    auto new_ws = editable->addWorksheet("NewSheet");
    ASSERT_NE(new_ws, nullptr);
    new_ws->writeString(0, 0, "New Sheet Data");
    
    // 检查未保存的更改
    EXPECT_TRUE(editable->hasUnsavedChanges());
    
    // 获取修改的工作表列表
    auto modified = editable->getModifiedWorksheets();
    EXPECT_GE(modified.size(), 1);
    
    // 保存更改
    ASSERT_TRUE(editable->save());
    
    // 验证更改已保存
    EXPECT_FALSE(editable->hasUnsavedChanges());
    
    // 重新打开验证
    auto verify = FastExcel::openForReading(core::Path(test_file_.string()));
    ASSERT_NE(verify, nullptr);
    
    EXPECT_EQ(verify->getWorksheetCount(), 2);
    
    auto ws_verify = verify->getWorksheet("Original");
    EXPECT_EQ(ws_verify->readString(0, 0), "Modified Data");
    EXPECT_DOUBLE_EQ(ws_verify->readNumber(1, 0), 200);
    EXPECT_EQ(ws_verify->readString(2, 0), "New Row");
    
    auto new_ws_verify = verify->getWorksheet("NewSheet");
    EXPECT_EQ(new_ws_verify->readString(0, 0), "New Sheet Data");
}

// 测试4：RowWriter 流式写入
TEST_F(NewArchitectureTest, RowWriterStreaming) {
    auto workbook = FastExcel::createWorkbook(core::Path(large_file_.string()));
    ASSERT_NE(workbook, nullptr);
    
    // 创建行写入器
    auto writer = workbook->createRowWriter("LargeData");
    ASSERT_NE(writer, nullptr);
    
    // 启用流式模式
    writer->enableStreamingMode();
    
    // 写入表头
    std::vector<std::string> headers = {"ID", "Name", "Value", "Status", "Date"};
    writer->writeHeader(headers);
    
    // 写入大量数据
    const int num_rows = 10000;
    auto start = std::chrono::high_resolution_clock::now();
    
    for (int i = 1; i <= num_rows; ++i) {
        writer->writeNumber(static_cast<double>(i))
              .writeString("Item_" + std::to_string(i))
              .writeNumber(i * 1.5)
              .writeBool(i % 2 == 0)
              .writeDateTime(std::chrono::system_clock::now())
              .nextRow();
        
        // 每1000行输出进度
        if (i % 1000 == 0) {
            auto stats = writer->getStats();
            EXPECT_EQ(stats.rows_written, i);
            EXPECT_TRUE(stats.is_streaming);
        }
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    // 写入汇总行
    writer->writeSummaryRow("Total", {
        "=COUNTA(A:A)-1",  // 计数
        "",
        "=SUM(C:C)",       // 求和
        "",
        ""
    });
    
    // 刷新并保存
    writer->flush();
    ASSERT_TRUE(workbook->save());
    
    // 输出性能信息
    std::cout << "Written " << num_rows << " rows in " 
              << duration.count() << "ms" << std::endl;
    
    // 验证文件
    auto verify = FastExcel::openForReading(core::Path(large_file_.string()));
    auto ws = verify->getWorksheet("LargeData");
    
    // 验证行数（包括表头和汇总行）
    EXPECT_EQ(ws->getRowCount(), num_rows + 2);
    
    // 验证一些数据
    EXPECT_EQ(ws->readString(0, 1), "Name");
    EXPECT_EQ(ws->readString(1, 1), "Item_1");
    EXPECT_DOUBLE_EQ(ws->readNumber(1, 2), 1.5);
    EXPECT_EQ(ws->readString(100, 1), "Item_100");
    EXPECT_DOUBLE_EQ(ws->readNumber(100, 2), 150.0);
}

// 测试5：缓存性能测试
TEST_F(NewArchitectureTest, CachePerformance) {
    // 创建测试文件
    {
        auto workbook = FastExcel::createWorkbook(core::Path(test_file_.string()));
        auto ws = workbook->addWorksheet("CacheTest");
        
        // 写入测试数据
        for (int i = 0; i < 100; ++i) {
            for (int j = 0; j < 10; ++j) {
                ws->writeNumber(i, j, i * 10 + j);
            }
        }
        
        ASSERT_TRUE(workbook->save());
    }
    
    // 只读模式打开
    auto readonly = FastExcel::openForReading(core::Path(test_file_.string()));
    auto ws = readonly->getWorksheet("CacheTest");
    
    // 第一次读取（缓存未命中）
    auto start1 = std::chrono::high_resolution_clock::now();
    for (int i = 0; i < 100; ++i) {
        for (int j = 0; j < 10; ++j) {
            auto value = ws->readNumber(i, j);
            EXPECT_DOUBLE_EQ(value, i * 10 + j);
        }
    }
    auto end1 = std::chrono::high_resolution_clock::now();
    auto duration1 = std::chrono::duration_cast<std::chrono::microseconds>(end1 - start1);
    
    // 第二次读取（缓存命中）
    auto start2 = std::chrono::high_resolution_clock::now();
    for (int i = 0; i < 100; ++i) {
        for (int j = 0; j < 10; ++j) {
            auto value = ws->readNumber(i, j);
            EXPECT_DOUBLE_EQ(value, i * 10 + j);
        }
    }
    auto end2 = std::chrono::high_resolution_clock::now();
    auto duration2 = std::chrono::duration_cast<std::chrono::microseconds>(end2 - start2);
    
    // 获取缓存统计
    auto stats = readonly->getCacheStats();
    
    // 验证缓存效果
    EXPECT_GT(stats.hits, 0);
    EXPECT_GT(stats.hit_rate, 0.0);
    
    // 第二次读取应该更快（由于缓存）
    std::cout << "First read: " << duration1.count() << "μs" << std::endl;
    std::cout << "Second read: " << duration2.count() << "μs" << std::endl;
    std::cout << "Cache hit rate: " << (stats.hit_rate * 100) << "%" << std::endl;
    
    // 通常缓存读取应该至少快2倍
    EXPECT_LT(duration2.count(), duration1.count());
}

// 测试6：从只读转换到编辑模式
TEST_F(NewArchitectureTest, ReadToEditTransition) {
    // 创建原始文件
    {
        auto workbook = FastExcel::createWorkbook(core::Path(test_file_.string()));
        auto ws = workbook->addWorksheet("Data");
        ws->writeString(0, 0, "Original");
        ASSERT_TRUE(workbook->save());
    }
    
    // 先以只读模式打开
    auto readonly = FastExcel::openForReading(core::Path(test_file_.string()));
    ASSERT_NE(readonly, nullptr);
    
    // 读取数据
    auto ws_read = readonly->getWorksheet("Data");
    EXPECT_EQ(ws_read->readString(0, 0), "Original");
    
    // 决定需要编辑，创建编辑会话
    fs::path edit_file = test_dir_ / "edited_copy.xlsx";
    auto editable = FastExcel::beginEdit(*readonly, core::Path(edit_file.string()));
    ASSERT_NE(editable, nullptr);
    
    // 编辑数据
    auto ws_edit = editable->getWorksheetForEdit("Data");
    ws_edit->writeString(0, 0, "Edited");
    ws_edit->writeString(1, 0, "New Data");
    
    // 保存到新文件
    ASSERT_TRUE(editable->save());
    
    // 验证原文件未改变
    auto verify_original = FastExcel::openForReading(core::Path(test_file_.string()));
    auto ws_original = verify_original->getWorksheet("Data");
    EXPECT_EQ(ws_original->readString(0, 0), "Original");
    
    // 验证新文件
    auto verify_new = FastExcel::openForReading(core::Path(edit_file.string()));
    auto ws_new = verify_new->getWorksheet("Data");
    EXPECT_EQ(ws_new->readString(0, 0), "Edited");
    EXPECT_EQ(ws_new->readString(1, 0), "New Data");
}

// 测试7：自动保存功能
TEST_F(NewArchitectureTest, AutoSaveFeature) {
    auto workbook = FastExcel::createWorkbook(core::Path(test_file_.string()));
    auto ws = workbook->addWorksheet("AutoSaveTest");
    
    // 启用自动保存（每2秒）
    workbook->enableAutoSave(std::chrono::seconds(2));
    
    // 写入数据
    ws->writeString(0, 0, "Auto Save Test");
    
    // 等待自动保存触发
    std::this_thread::sleep_for(std::chrono::seconds(3));
    
    // 获取统计信息
    auto stats = workbook->getStats();
    EXPECT_GT(stats.save_count, 0);
    EXPECT_TRUE(stats.auto_save_enabled);
    
    // 禁用自动保存
    workbook->disableAutoSave();
    
    // 验证文件已保存
    ASSERT_TRUE(fs::exists(test_file_));
}

// 测试8：文件验证功能
TEST_F(NewArchitectureTest, FileValidation) {
    // 测试有效的Excel文件
    {
        auto workbook = FastExcel::createWorkbook(core::Path(test_file_.string()));
        auto ws = workbook->addWorksheet("Valid");
        ws->writeString(0, 0, "Valid File");
        ASSERT_TRUE(workbook->save());
    }
    
    EXPECT_TRUE(FastExcel::isValidExcelFile(core::Path(test_file_.string())));
    
    // 测试无效文件
    fs::path invalid_file = test_dir_ / "invalid.xlsx";
    std::ofstream ofs(invalid_file);
    ofs << "This is not an Excel file";
    ofs.close();
    
    EXPECT_FALSE(FastExcel::isValidExcelFile(core::Path(invalid_file.string())));
    
    // 测试不存在的文件
    fs::path nonexistent = test_dir_ / "nonexistent.xlsx";
    EXPECT_FALSE(FastExcel::isValidExcelFile(core::Path(nonexistent.string())));
}

// 测试9：获取文件信息
TEST_F(NewArchitectureTest, GetFileInfo) {
    // 创建测试文件
    {
        auto workbook = FastExcel::createWorkbook(core::Path(test_file_.string()));
        workbook->addWorksheet("Sheet1");
        workbook->addWorksheet("Sheet2");
        workbook->addWorksheet("Data");
        
        auto ws = workbook->getWorksheetForEdit("Data");
        for (int i = 0; i < 100; ++i) {
            ws->writeNumber(i, 0, i);
        }
        
        ASSERT_TRUE(workbook->save());
    }
    
    // 获取文件信息
    auto info = FastExcel::getFileInfo(core::Path(test_file_.string()));
    
    EXPECT_TRUE(info.is_valid);
    EXPECT_EQ(info.worksheet_names.size(), 3);
    EXPECT_EQ(info.worksheet_names[0], "Sheet1");
    EXPECT_EQ(info.worksheet_names[1], "Sheet2");
    EXPECT_EQ(info.worksheet_names[2], "Data");
    EXPECT_GT(info.estimated_size, 0);
}

// 测试10：批量转换功能
TEST_F(NewArchitectureTest, BatchConversion) {
    // 创建多个源文件
    std::vector<core::Path> input_files;
    for (int i = 1; i <= 3; ++i) {
        fs::path file = test_dir_ / ("source" + std::to_string(i) + ".xlsx");
        auto workbook = FastExcel::createWorkbook(core::Path(file.string()));
        auto ws = workbook->addWorksheet("Data");
        ws->writeString(0, 0, "File " + std::to_string(i));
        ws->writeNumber(1, 0, i * 100);
        ASSERT_TRUE(workbook->save());
        input_files.push_back(core::Path(file.string()));
    }
    
    // 定义转换函数（添加新列）
    auto converter = [](const read::ReadWorkbook& source, edit::EditSession& target) {
        auto src_ws = source.getWorksheet("Data");
        auto tgt_ws = target.addWorksheet("Converted");
        
        // 复制原始数据并添加新列
        tgt_ws->writeString(0, 0, src_ws->readString(0, 0));
        tgt_ws->writeNumber(1, 0, src_ws->readNumber(1, 0));
        tgt_ws->writeString(0, 1, "Converted");
        tgt_ws->writeNumber(1, 1, src_ws->readNumber(1, 0) * 2);
        
        return true;
    };
    
    // 执行批量转换
    fs::path output_dir = test_dir_ / "output";
    fs::create_directories(output_dir);
    
    int converted = FastExcel::batchConvert(
        input_files, 
        core::Path(output_dir.string()),
        converter
    );
    
    EXPECT_EQ(converted, 3);
    
    // 验证转换后的文件
    for (int i = 1; i <= 3; ++i) {
        fs::path converted_file = output_dir / ("source" + std::to_string(i) + "_converted.xlsx");
        ASSERT_TRUE(fs::exists(converted_file));
        
        auto verify = FastExcel::openForReading(core::Path(converted_file.string()));
        auto ws = verify->getWorksheet("Converted");
        
        EXPECT_EQ(ws->readString(0, 0), "File " + std::to_string(i));
        EXPECT_DOUBLE_EQ(ws->readNumber(1, 0), i * 100);
        EXPECT_EQ(ws->readString(0, 1), "Converted");
        EXPECT_DOUBLE_EQ(ws->readNumber(1, 1), i * 200);
    }
}

// 主函数
int main(int argc, char** argv) {
    ::testing::InitGoogleTest(&argc, argv);
    return RUN_ALL_TESTS();
}