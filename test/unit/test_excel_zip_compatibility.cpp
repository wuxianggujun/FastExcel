/**
 * @file test_excel_zip_compatibility.cpp
 * @brief 测试ZIP文件与Excel的兼容性
 * 
 * 这个测试专门验证生成的XLSX文件是否能被Excel正常打开
 */

#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include "fastexcel/FastExcel.hpp"
#include <gtest/gtest.h>
#include <filesystem>
#include <fstream>
#include <vector>
#include <chrono>
#include <ctime>

namespace fastexcel {
namespace archive {

class ExcelZipCompatibilityTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化日志系统
        fastexcel::Logger::getInstance().initialize("logs/excel_zip_compatibility_test.log",
                                                   fastexcel::Logger::Level::DEBUG,
                                                   false);
        
        // 创建测试目录
        test_dir_ = "test_excel_compatibility";
        std::filesystem::create_directories(test_dir_);
        
        // 为每个测试创建唯一的文件路径
        static int test_counter = 0;
        test_file_prefix_ = test_dir_ + "/excel_test_" + std::to_string(++test_counter);
    }

    void TearDown() override {
        // 清理测试文件
        try {
            // 不自动删除文件，以便手动检查
            // 如果需要自动清理，取消下面的注释
            // std::filesystem::remove_all(test_dir_);
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_WARN("Failed to clean up test directory: {}", e.what());
        }
        
        // 关闭日志系统
        fastexcel::Logger::getInstance().shutdown();
    }

    // 辅助函数：创建最小的Excel XML结构
    void createMinimalExcelStructure(const std::string& filename) {
        ZipArchive zip(filename);
        ASSERT_TRUE(zip.open(true));
        
        // [Content_Types].xml
        std::string content_types = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>)";
        EXPECT_EQ(zip.addFile("[Content_Types].xml", content_types), ZipError::Ok);
        
        // _rels/.rels
        std::string root_rels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>)";
        EXPECT_EQ(zip.addFile("_rels/.rels", root_rels), ZipError::Ok);
        
        // docProps/app.xml
        std::string app_props = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>1</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>Sheet1</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company>FastExcel</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>)";
        EXPECT_EQ(zip.addFile("docProps/app.xml", app_props), ZipError::Ok);
        
        // docProps/core.xml - 使用封装的TimeUtils类
        std::tm current_time = utils::TimeUtils::getCurrentUTCTime();
        std::string iso_time = utils::TimeUtils::formatTimeISO8601(current_time);
        
        std::string core_props = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>FastExcel Test</dc:creator>
  <cp:lastModifiedBy>FastExcel Test</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">)" + iso_time + R"(</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">)" + iso_time + R"(</dcterms:modified>
</cp:coreProperties>)";
        EXPECT_EQ(zip.addFile("docProps/core.xml", core_props), ZipError::Ok);
        
        // xl/_rels/workbook.xml.rels
        std::string workbook_rels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>)";
        EXPECT_EQ(zip.addFile("xl/_rels/workbook.xml.rels", workbook_rels), ZipError::Ok);
        
        // xl/workbook.xml
        std::string workbook = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
  <workbookPr defaultThemeVersion="124226"/>
  <bookViews>
    <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>
  </bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>)";
        EXPECT_EQ(zip.addFile("xl/workbook.xml", workbook), ZipError::Ok);
        
        // xl/styles.xml (minimal)
        std::string styles = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>)";
        EXPECT_EQ(zip.addFile("xl/styles.xml", styles), ZipError::Ok);
        
        // xl/theme/theme1.xml (minimal)
        std::string theme = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
            <a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
            <a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>)";
        EXPECT_EQ(zip.addFile("xl/theme/theme1.xml", theme), ZipError::Ok);
        
        // xl/worksheets/sheet1.xml
        std::string worksheet = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B2"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:2">
      <c r="A1" t="inlineStr">
        <is><t>Test</t></is>
      </c>
      <c r="B1">
        <v>123</v>
      </c>
    </row>
    <row r="2" spans="1:2">
      <c r="A2" t="inlineStr">
        <is><t>Excel Compatibility</t></is>
      </c>
      <c r="B2">
        <v>456.78</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>)";
        EXPECT_EQ(zip.addFile("xl/worksheets/sheet1.xml", worksheet), ZipError::Ok);
        
        ASSERT_TRUE(zip.close());
    }

    std::string test_dir_;
    std::string test_file_prefix_;
};

// 测试1: 创建最小的Excel兼容文件
TEST_F(ExcelZipCompatibilityTest, CreateMinimalExcelFile) {
    FASTEXCEL_LOG_INFO("Testing minimal Excel file creation");
    
    std::string filename = test_file_prefix_ + "_minimal.xlsx";
    createMinimalExcelStructure(filename);
    
    // 验证文件存在
    EXPECT_TRUE(std::filesystem::exists(filename));
    
    // 验证文件大小合理
    auto file_size = std::filesystem::file_size(filename);
    EXPECT_GT(file_size, 1000); // 至少1KB
    EXPECT_LT(file_size, 100000); // 不超过100KB
    
    FASTEXCEL_LOG_INFO("Created minimal Excel file: {}", filename);
    FASTEXCEL_LOG_INFO("File size: {} bytes", file_size);
    FASTEXCEL_LOG_INFO("Please open this file in Excel to verify compatibility");
}

// 测试2: 使用FastExcel API创建文件
TEST_F(ExcelZipCompatibilityTest, CreateWithFastExcelAPI) {
    FASTEXCEL_LOG_INFO("Testing Excel file creation with FastExcel API");
    
    std::string filename = test_file_prefix_ + "_api.xlsx";
    
    // 初始化FastExcel - 使用void版本
    void (*init_func)() = fastexcel::initialize;
    init_func();
    
    // 创建工作簿
    auto workbook = core::Workbook::create(filename);
    ASSERT_TRUE(workbook->open());
    
    // 添加工作表
    auto worksheet = workbook->addWorksheet("TestSheet");
    ASSERT_NE(worksheet, nullptr);
    
    // 写入测试数据
    worksheet->writeString(0, 0, "ZIP Fix Test");
    worksheet->writeNumber(0, 1, 2025);
    worksheet->writeString(1, 0, "Version Info");
    worksheet->writeNumber(1, 1, 2580); // Windows version_madeby
    worksheet->writeString(2, 0, "Compression");
    worksheet->writeString(2, 1, "STORE");
    
    // 设置文档属性
    workbook->setTitle("Excel ZIP Compatibility Test");
    workbook->setAuthor("FastExcel");
    workbook->setSubject("Testing ZIP format fixes");
    
    // 保存文件
    ASSERT_TRUE(workbook->save());
    workbook->close();
    
    // 清理
    fastexcel::cleanup();
    
    // 验证文件
    EXPECT_TRUE(std::filesystem::exists(filename));
    auto file_size = std::filesystem::file_size(filename);
    FASTEXCEL_LOG_INFO("Created Excel file with API: {}", filename);
    FASTEXCEL_LOG_INFO("File size: {} bytes", file_size);
}

// 测试3: 批量文件写入
TEST_F(ExcelZipCompatibilityTest, BatchFileWriting) {
    FASTEXCEL_LOG_INFO("Testing batch file writing");
    
    std::string filename = test_file_prefix_ + "_batch.xlsx";
    ZipArchive zip(filename);
    ASSERT_TRUE(zip.open(true));
    
    // 准备批量写入的文件
    std::vector<ZipArchive::FileEntry> files;
    
    // 添加所有必要的Excel文件
    files.emplace_back("[Content_Types].xml", 
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>)");
    
    // 使用批量写入
    EXPECT_EQ(zip.addFiles(files), ZipError::Ok);
    
    ASSERT_TRUE(zip.close());
    
    // 验证文件
    EXPECT_TRUE(std::filesystem::exists(filename));
    FASTEXCEL_LOG_INFO("Created Excel file with batch writing: {}", filename);
}

// 测试4: 流式写入
TEST_F(ExcelZipCompatibilityTest, StreamingFileWriting) {
    FASTEXCEL_LOG_INFO("Testing streaming file writing");
    
    std::string filename = test_file_prefix_ + "_streaming.xlsx";
    ZipArchive zip(filename);
    ASSERT_TRUE(zip.open(true));
    
    // 使用流式写入创建worksheet
    std::string worksheet_path = "xl/worksheets/sheet1.xml";
    EXPECT_EQ(zip.openEntry(worksheet_path), ZipError::Ok);
    
    // 写入XML头
    std::string xml_header = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B100"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>)";
    EXPECT_EQ(zip.writeChunk(xml_header.data(), xml_header.size()), ZipError::Ok);
    
    // 流式写入100行数据
    for (int i = 1; i <= 100; ++i) {
        std::string row = R"(
    <row r=")" + std::to_string(i) + R"(" spans="1:2">
      <c r="A)" + std::to_string(i) + R"(" t="inlineStr">
        <is><t>Row )" + std::to_string(i) + R"(</t></is>
      </c>
      <c r="B)" + std::to_string(i) + R"(">
        <v>)" + std::to_string(i * 100) + R"(</v>
      </c>
    </row>)";
        EXPECT_EQ(zip.writeChunk(row.data(), row.size()), ZipError::Ok);
    }
    
    // 写入XML尾
    std::string xml_footer = R"(
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>)";
    EXPECT_EQ(zip.writeChunk(xml_footer.data(), xml_footer.size()), ZipError::Ok);
    
    EXPECT_EQ(zip.closeEntry(), ZipError::Ok);
    
    // 添加其他必要文件（简化）
    EXPECT_EQ(zip.addFile("[Content_Types].xml", "..."), ZipError::Ok);
    
    ASSERT_TRUE(zip.close());
    
    // 验证文件
    EXPECT_TRUE(std::filesystem::exists(filename));
    FASTEXCEL_LOG_INFO("Created Excel file with streaming: {}", filename);
}

// 测试5: 验证修复后的设置
TEST_F(ExcelZipCompatibilityTest, VerifyFixedSettings) {
    FASTEXCEL_LOG_INFO("Testing fixed ZIP settings");
    
    std::string filename = test_file_prefix_ + "_verify.xlsx";
    
    // 创建一个简单的Excel文件
    ZipArchive zip(filename);
    ASSERT_TRUE(zip.open(true));
    
    // 添加一个测试文件
    std::string test_content = "Test content for ZIP settings verification";
    EXPECT_EQ(zip.addFile("test.txt", test_content), ZipError::Ok);
    
    ASSERT_TRUE(zip.close());
    
    // 重新打开并读取
    ASSERT_TRUE(zip.open(false));
    
    std::string extracted;
    EXPECT_EQ(zip.extractFile("test.txt", extracted), ZipError::Ok);
    EXPECT_EQ(extracted, test_content);
    
    zip.close();
    
    FASTEXCEL_LOG_INFO("ZIP settings verification completed");
    FASTEXCEL_LOG_INFO("File created: {}", filename);
    
    // 注意：实际的version_madeby和压缩方法验证需要使用外部工具
    // 如 7-Zip 或专门的ZIP分析工具
    FASTEXCEL_LOG_INFO("To verify ZIP metadata, use: 7z l -slt {}", filename);
}

}} // namespace fastexcel::archive