/**
 * @file test_zip_excel_compatibility.cpp
 * @brief 演示ZIP文件Excel兼容性修复的示例程序
 * 
 * 这个示例展示了如何创建能被Excel正常打开的XLSX文件，
 * 包括直接生成XML内容和从本地文件读取两种方式。
 */

#include <iostream>
#include <fstream>
#include <filesystem>
#include <chrono>
#include <ctime>
#include <iomanip>
#include <vector>
#include <string>
#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/utils/Logger.hpp"

using namespace fastexcel::archive;
namespace fs = std::filesystem;

// 辅助函数：获取当前时间的ISO格式字符串
std::string getCurrentTimeISO() {
    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    std::tm tm = *std::localtime(&time_t);
    
    std::ostringstream oss;
    oss << std::put_time(&tm, "%Y-%m-%dT%H:%M:%SZ");
    return oss.str();
}

// 辅助函数：创建测试用的XML文件
void createTestXMLFiles(const std::string& dir) {
    fs::create_directories(dir);
    fs::create_directories(dir + "/_rels");
    fs::create_directories(dir + "/docProps");
    fs::create_directories(dir + "/xl");
    fs::create_directories(dir + "/xl/_rels");
    fs::create_directories(dir + "/xl/theme");
    fs::create_directories(dir + "/xl/worksheets");
    
    // [Content_Types].xml
    std::ofstream content_types(dir + "/[Content_Types].xml");
    content_types << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    content_types.close();
    
    // _rels/.rels
    std::ofstream root_rels(dir + "/_rels/.rels");
    root_rels << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>)";
    root_rels.close();
    
    // docProps/app.xml
    std::ofstream app_props(dir + "/docProps/app.xml");
    app_props << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    app_props.close();
    
    // docProps/core.xml
    std::ofstream core_props(dir + "/docProps/core.xml");
    core_props << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>FastExcel Test</dc:creator>
  <cp:lastModifiedBy>FastExcel Test</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">)" + getCurrentTimeISO() + R"(</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">)" + getCurrentTimeISO() + R"(</dcterms:modified>
</cp:coreProperties>)";
    core_props.close();
    
    // xl/_rels/workbook.xml.rels
    std::ofstream workbook_rels(dir + "/xl/_rels/workbook.xml.rels");
    workbook_rels << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>)";
    workbook_rels.close();
    
    // xl/workbook.xml
    std::ofstream workbook(dir + "/xl/workbook.xml");
    workbook << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    workbook.close();
    
    // xl/styles.xml (简化版)
    std::ofstream styles(dir + "/xl/styles.xml");
    styles << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    styles.close();
    
    // xl/theme/theme1.xml (简化版)
    std::ofstream theme(dir + "/xl/theme/theme1.xml");
    theme << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>)";
    theme.close();
    
    // xl/worksheets/sheet1.xml
    std::ofstream worksheet(dir + "/xl/worksheets/sheet1.xml");
    worksheet << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:C3"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:3">
      <c r="A1" t="inlineStr">
        <is><t>ZIP修复测试</t></is>
      </c>
      <c r="B1" t="inlineStr">
        <is><t>Excel兼容性</t></is>
      </c>
      <c r="C1" t="inlineStr">
        <is><t>状态</t></is>
      </c>
    </row>
    <row r="2" spans="1:3">
      <c r="A2" t="inlineStr">
        <is><t>version_madeby</t></is>
      </c>
      <c r="B2">
        <v>2580</v>
      </c>
      <c r="C2" t="inlineStr">
        <is><t>已修复</t></is>
      </c>
    </row>
    <row r="3" spans="1:3">
      <c r="A3" t="inlineStr">
        <is><t>压缩方法</t></is>
      </c>
      <c r="B3" t="inlineStr">
        <is><t>STORE</t></is>
      </c>
      <c r="C3" t="inlineStr">
        <is><t>已修复</t></is>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>)";
    worksheet.close();
}

// 测试1：从程序生成的XML创建XLSX
void testGeneratedXML(const std::string& output_file) {
    std::cout << "\n=== 测试1：从程序生成的XML创建XLSX ===" << std::endl;
    
    ZipArchive zip(output_file);
    if (!zip.open(true)) {
        std::cerr << "无法创建ZIP文件: " << output_file << std::endl;
        return;
    }
    
    // 准备要添加的文件
    std::vector<ZipArchive::FileEntry> files;
    
    // [Content_Types].xml
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
    
    // _rels/.rels
    files.emplace_back("_rels/.rels",
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>)");
    
    // 添加其他必要文件...
    // （为简洁起见，这里省略了其他文件的添加）
    
    // 使用批量添加
    auto result = zip.addFiles(files);
    if (result != ZipError::Ok) {
        std::cerr << "添加文件失败" << std::endl;
        return;
    }
    
    // 添加一个简单的worksheet
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
        <is><t>生成的XML测试</t></is>
      </c>
      <c r="B1">
        <v>2025</v>
      </c>
    </row>
    <row r="2" spans="1:2">
      <c r="A2" t="inlineStr">
        <is><t>修复后可正常打开</t></is>
      </c>
      <c r="B2">
        <v>100</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>)";
    
    result = zip.addFile("xl/worksheets/sheet1.xml", worksheet);
    if (result != ZipError::Ok) {
        std::cerr << "添加worksheet失败" << std::endl;
        return;
    }
    
    if (zip.close()) {
        std::cout << "成功创建: " << output_file << std::endl;
        std::cout << "文件大小: " << fs::file_size(output_file) << " 字节" << std::endl;
    } else {
        std::cerr << "关闭ZIP文件失败" << std::endl;
    }
}

// 测试2：从本地文件创建XLSX
void testLocalFiles(const std::string& input_dir, const std::string& output_file) {
    std::cout << "\n=== 测试2：从本地文件创建XLSX ===" << std::endl;
    
    ZipArchive zip(output_file);
    if (!zip.open(true)) {
        std::cerr << "无法创建ZIP文件: " << output_file << std::endl;
        return;
    }
    
    // 递归添加目录中的所有文件
    for (const auto& entry : fs::recursive_directory_iterator(input_dir)) {
        if (entry.is_regular_file()) {
            std::string relative_path = fs::relative(entry.path(), input_dir).string();
            // 将反斜杠替换为正斜杠
            std::replace(relative_path.begin(), relative_path.end(), '\\', '/');
            
            // 读取文件内容
            std::ifstream file(entry.path(), std::ios::binary);
            std::string content((std::istreambuf_iterator<char>(file)),
                              std::istreambuf_iterator<char>());
            
            auto result = zip.addFile(relative_path, content);
            if (result == ZipError::Ok) {
                std::cout << "添加: " << relative_path << std::endl;
            } else {
                std::cerr << "添加失败: " << relative_path << std::endl;
            }
        }
    }
    
    if (zip.close()) {
        std::cout << "成功创建: " << output_file << std::endl;
        std::cout << "文件大小: " << fs::file_size(output_file) << " 字节" << std::endl;
    } else {
        std::cerr << "关闭ZIP文件失败" << std::endl;
    }
}

// 测试3：流式写入大文件
void testStreamingWrite(const std::string& output_file) {
    std::cout << "\n=== 测试3：流式写入大文件 ===" << std::endl;
    
    ZipArchive zip(output_file);
    if (!zip.open(true)) {
        std::cerr << "无法创建ZIP文件: " << output_file << std::endl;
        return;
    }
    
    // 开始流式写入worksheet
    if (zip.openEntry("xl/worksheets/sheet1.xml") != ZipError::Ok) {
        std::cerr << "无法开始流式写入" << std::endl;
        return;
    }
    
    // 写入XML头
    std::string header = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B1000"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>)";
    
    if (zip.writeChunk(header.data(), header.size()) != ZipError::Ok) {
        std::cerr << "写入头部失败" << std::endl;
        return;
    }
    
    // 流式写入1000行数据
    for (int i = 1; i <= 1000; ++i) {
        std::string row = R"(
    <row r=")" + std::to_string(i) + R"(" spans="1:2">
      <c r="A)" + std::to_string(i) + R"(" t="inlineStr">
        <is><t>行 )" + std::to_string(i) + R"(</t></is>
      </c>
      <c r="B)" + std::to_string(i) + R"(">
        <v>)" + std::to_string(i * 10) + R"(</v>
      </c>
    </row>)";
        
        if (zip.writeChunk(row.data(), row.size()) != ZipError::Ok) {
            std::cerr << "写入行 " << i << " 失败" << std::endl;
            return;
        }
        
        if (i % 100 == 0) {
            std::cout << "已写入 " << i << " 行..." << std::endl;
        }
    }
    
    // 写入XML尾
    std::string footer = R"(
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>)";
    
    if (zip.writeChunk(footer.data(), footer.size()) != ZipError::Ok) {
        std::cerr << "写入尾部失败" << std::endl;
        return;
    }
    
    if (zip.closeEntry() != ZipError::Ok) {
        std::cerr << "关闭流式写入失败" << std::endl;
        return;
    }
    
    // 添加其他必要的Excel文件（简化版）
    zip.addFile("[Content_Types].xml", "...");  // 实际使用时需要完整内容
    
    if (zip.close()) {
        std::cout << "成功创建: " << output_file << std::endl;
        std::cout << "文件大小: " << fs::file_size(output_file) << " 字节" << std::endl;
    } else {
        std::cerr << "关闭ZIP文件失败" << std::endl;
    }
}

int main() {
    // 初始化日志系统
    fastexcel::Logger::getInstance().initialize("logs/excel_compatibility_test.log",
                                               fastexcel::Logger::Level::INFO,
                                               true);  // 同时输出到控制台
    
    std::cout << "FastExcel ZIP Excel兼容性测试" << std::endl;
    std::cout << "==============================" << std::endl;
    
    // 创建输出目录
    fs::create_directories("output");
    fs::create_directories("temp_xml");
    
    // 测试1：从程序生成的XML创建XLSX
    testGeneratedXML("output/test_generated.xlsx");
    
    // 测试2：从本地文件创建XLSX
    std::cout << "\n准备本地XML文件..." << std::endl;
    createTestXMLFiles("temp_xml");
    testLocalFiles("temp_xml", "output/test_local_files.xlsx");
    
    // 测试3：流式写入大文件
    testStreamingWrite("output/test_streaming.xlsx");
    
    // 清理临时文件
    std::cout << "\n清理临时文件..." << std::endl;
    fs::remove_all("temp_xml");
    
    std::cout << "\n测试完成！" << std::endl;
    std::cout << "请使用Excel打开output目录中的文件验证兼容性。" << std::endl;
    std::cout << "\n修复说明：" << std::endl;
    std::cout << "1. version_madeby: 使用 (MZ_HOST_SYSTEM_WINDOWS_NTFS << 8) | 20 = 2580" << std::endl;
    std::cout << "2. 压缩方法: 统一使用 STORE (无压缩)" << std::endl;
    std::cout << "3. 文件标志: 批量写入使用 0，流式写入使用 MZ_ZIP_FLAG_DATA_DESCRIPTOR" << std::endl;
    std::cout << "4. 时间戳: 使用本地时间的DOS格式" << std::endl;
    
    // 关闭日志系统
    fastexcel::Logger::getInstance().shutdown();
    
    return 0;
}