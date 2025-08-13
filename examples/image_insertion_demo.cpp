/**
 * @file image_insertion_demo.cpp
 * @brief FastExcel图片插入功能演示代码
 *
 * 本示例展示了FastExcel库的图片插入功能：
 * 1. 基本图片插入到单元格
 * 2. 图片插入到指定范围
 * 3. 绝对定位图片插入
 * 4. 使用Excel地址格式插入图片
 * 5. 批量图片管理
 * 6. 图片属性设置
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/Image.hpp"
#include "fastexcel/xml/DrawingXMLGenerator.hpp"

#include <iostream>
#include <vector>
#include <memory>

using namespace fastexcel;
using namespace fastexcel::core;

void demonstrateBasicImageInsertion() {
    std::cout << "\n=== 1. 基本图片插入演示 ===" << std::endl;
    
    // 创建工作簿
    auto workbook = Workbook::create(Path("images_basic.xlsx"));
    if (!workbook) {
        std::cerr << "无法创建工作簿" << std::endl;
        return;
    }
    
    auto worksheet = workbook->addSheet("基本图片");
    
    // 添加一些标题
    worksheet->setValue(0, 0, std::string("图片插入演示"));
    worksheet->setValue(2, 0, std::string("单元格锚定图片:"));
    worksheet->setValue(8, 0, std::string("范围锚定图片:"));
    
    // 方式1：插入图片到单元格（使用真实的图片文件）
    std::string image_path = "tinaimage.png";
    
    std::cout << "尝试插入图片: " << image_path << std::endl;
    
    // 从文件加载图片
    auto image = Image::fromFile(image_path);
    if (image) {
        image->setName("ChatGPT示例图片");
        image->setDescription("基本单元格锚定图片");
        
        std::cout << "成功加载图片: " << image->getName() << std::endl;
        std::cout << "图片尺寸: " << image->getOriginalWidth() << "x" << image->getOriginalHeight() << std::endl;
        
        std::string image_id = worksheet->insertImage(2, 1, std::move(image));
        if (!image_id.empty()) {
            std::cout << "成功插入图片到B3单元格，ID: " << image_id << std::endl;
        } else {
            std::cout << "图片插入失败" << std::endl;
        }
    } else {
        std::cout << "无法加载图片文件: " << image_path << std::endl;
        std::cout << "请确保图片文件存在" << std::endl;
    }
    
    // 保存工作簿
    if (workbook->save()) {
        std::cout << "工作簿保存成功: images_basic.xlsx" << std::endl;
    } else {
        std::cout << "工作簿保存失败" << std::endl;
    }
}

void demonstrateAdvancedImageInsertion() {
    std::cout << "\n=== 2. 高级图片插入演示 ===" << std::endl;
    
    auto workbook = Workbook::create(Path("images_advanced.xlsx"));
    if (!workbook) {
        std::cerr << "无法创建工作簿" << std::endl;
        return;
    }
    
    auto worksheet = workbook->addSheet("高级图片");
    
    // 添加标题
    worksheet->setValue(0, 0, std::string("高级图片插入演示"));
    
    // 使用真实的图片文件进行各种插入方式演示
    std::string image_path = "tinaimage.png";
    
    // 方式1：范围锚定图片
    worksheet->setValue(2, 0, std::string("范围锚定图片 (A3:C5):"));
    auto range_image = Image::fromFile(image_path);
    if (range_image) {
        range_image->setName("范围锚定图片");
        range_image->setDescription("锚定到A3:C5范围的图片");
        
        std::string image_id = worksheet->insertImage(2, 0, 4, 2, std::move(range_image));
        if (!image_id.empty()) {
            std::cout << "成功插入范围锚定图片到A3:C5，ID: " << image_id << std::endl;
        }
    }
    
    // 方式2：绝对定位图片
    worksheet->setValue(6, 0, std::string("绝对定位图片:"));
    auto absolute_image = Image::fromFile(image_path);
    if (absolute_image) {
        absolute_image->setName("绝对定位图片");
        absolute_image->setDescription("绝对定位在(300,200)的图片，尺寸150x120");
        
        std::string image_id = worksheet->insertImageAt(300, 200, 150, 120, std::move(absolute_image));
        if (!image_id.empty()) {
            std::cout << "成功插入绝对定位图片，ID: " << image_id << std::endl;
        }
    }
    
    // 方式3：使用Excel地址格式
    worksheet->setValue(8, 0, std::string("Excel地址格式图片:"));
    auto address_image = Image::fromFile(image_path);
    if (address_image) {
        address_image->setName("地址格式图片");
        address_image->setDescription("使用Excel地址B9插入的图片");
        
        std::string image_id = worksheet->insertImage("B9", std::move(address_image));
        if (!image_id.empty()) {
            std::cout << "成功插入地址格式图片到B9，ID: " << image_id << std::endl;
        }
    }
    
    // 方式4：使用Excel范围格式
    worksheet->setValue(10, 0, std::string("Excel范围格式图片:"));
    auto range_address_image = Image::fromFile(image_path);
    if (range_address_image) {
        range_address_image->setName("范围地址图片");
        range_address_image->setDescription("使用Excel范围D11:F13插入的图片");
        
        std::string image_id = worksheet->insertImageRange("D11:F13", std::move(range_address_image));
        if (!image_id.empty()) {
            std::cout << "成功插入范围地址格式图片到D11:F13，ID: " << image_id << std::endl;
        }
    }
    
    // 保存工作簿
    if (workbook->save()) {
        std::cout << "工作簿保存成功: images_advanced.xlsx" << std::endl;
    } else {
        std::cout << "工作簿保存失败" << std::endl;
    }
}

void demonstrateImageManagement() {
    std::cout << "\n=== 3. 图片管理演示 ===" << std::endl;
    
    auto workbook = Workbook::create(Path("images_management.xlsx"));
    if (!workbook) {
        std::cerr << "无法创建工作簿" << std::endl;
        return;
    }
    
    auto worksheet = workbook->addSheet("图片管理");
    
    // 添加标题
    worksheet->setValue(0, 0, std::string("图片管理演示"));
    
    // 使用真实图片文件进行批量插入演示
    std::string image_path = "tinaimage.png";
    
    // 插入多个图片
    std::vector<std::string> image_ids;
    
    for (int i = 0; i < 3; ++i) {
        auto image = Image::fromFile(image_path);
        if (image) {
            image->setName("管理测试图片 " + std::to_string(i + 1));
            image->setDescription("第 " + std::to_string(i + 1) + " 个用于管理演示的图片");
            
            std::string image_id = worksheet->insertImage(2 + i * 3, 1, std::move(image)); // 每隔3行插入
            if (!image_id.empty()) {
                image_ids.push_back(image_id);
                std::cout << "插入图片 " << (i + 1) << "，ID: " << image_id << std::endl;
            }
        }
    }
    
    // 显示图片统计信息
    std::cout << "工作表中的图片数量: " << worksheet->getImageCount() << std::endl;
    std::cout << "图片占用内存: " << worksheet->getImagesMemoryUsage() << " 字节" << std::endl;
    
    // 查找图片
    if (!image_ids.empty()) {
        const Image* found_image = worksheet->findImage(image_ids[0]);
        if (found_image) {
            std::cout << "找到图片: " << found_image->getName() 
                     << " (格式: " << static_cast<int>(found_image->getFormat()) << ")" << std::endl;
        }
    }
    
    // 删除一个图片
    if (image_ids.size() > 1) {
        bool removed = worksheet->removeImage(image_ids[1]);
        if (removed) {
            std::cout << "成功删除图片: " << image_ids[1] << std::endl;
            std::cout << "删除后图片数量: " << worksheet->getImageCount() << std::endl;
        }
    }
    
    // 保存工作簿
    if (workbook->save()) {
        std::cout << "工作簿保存成功: images_management.xlsx" << std::endl;
    } else {
        std::cout << "工作簿保存失败" << std::endl;
    }
}

void demonstrateImageFormats() {
    std::cout << "\n=== 4. 图片格式演示 ===" << std::endl;
    
    // 演示不同图片格式的支持
    std::cout << "支持的图片格式:" << std::endl;
    std::cout << "- PNG: " << core::ImageUtils::formatToString(ImageFormat::PNG) << std::endl;
    std::cout << "- JPEG: " << core::ImageUtils::formatToString(ImageFormat::JPEG) << std::endl;
    std::cout << "- GIF: " << core::ImageUtils::formatToString(ImageFormat::GIF) << std::endl;
    std::cout << "- BMP: " << core::ImageUtils::formatToString(ImageFormat::BMP) << std::endl;
    
    // 演示格式检测
    std::vector<uint8_t> png_signature = {0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A};
    auto detected_format = core::ImageUtils::formatFromExtension("test.png");
    std::cout << "从扩展名检测格式 'test.png': " << core::ImageUtils::formatToString(detected_format) << std::endl;
    
    // 演示MIME类型
    std::cout << "PNG MIME类型: " << core::ImageUtils::getMimeType(ImageFormat::PNG) << std::endl;
    std::cout << "JPEG MIME类型: " << core::ImageUtils::getMimeType(ImageFormat::JPEG) << std::endl;
}

int main() {
    std::cout << "FastExcel 图片插入功能演示程序" << std::endl;
    std::cout << "=================================" << std::endl;
    
    // 初始化FastExcel
    if (!fastexcel::initialize("logs/image_demo.log", true)) {
        std::cerr << "FastExcel初始化失败" << std::endl;
        return 1;
    }
    
    try {
        // 演示各种图片插入功能
        demonstrateBasicImageInsertion();
        demonstrateAdvancedImageInsertion();
        demonstrateImageManagement();
        demonstrateImageFormats();
        
        std::cout << "\n=== 演示完成 ===" << std::endl;
        std::cout << "生成的文件:" << std::endl;
        std::cout << "- images_basic.xlsx - 基本图片插入演示" << std::endl;
        std::cout << "- images_advanced.xlsx - 高级图片插入演示" << std::endl;
        std::cout << "- images_management.xlsx - 图片管理演示" << std::endl;
        std::cout << "\n注意：由于使用了内存中的简单PNG数据，生成的图片可能很小。" << std::endl;
        std::cout << "在实际使用中，请使用真实的图片文件。" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "演示过程中发生错误: " << e.what() << std::endl;
        fastexcel::cleanup();
        return 1;
    }
    
    // 清理资源
    fastexcel::cleanup();
    
    std::cout << "\n程序执行完成。请查看日志文件 logs/image_demo.log 获取详细信息。" << std::endl;
    return 0;
}