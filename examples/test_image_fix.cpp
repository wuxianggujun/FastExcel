#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/Image.hpp"
#include <iostream>
#include <fstream>

using namespace fastexcel;
using namespace fastexcel::core;

// 创建一个简单的测试PNG图片（10x10像素的红色图片）
std::vector<uint8_t> createTestPNG() {
    // 创建一个10x10的红色PNG图片
    // 这是一个有效的PNG文件，显示为红色正方形
    std::vector<uint8_t> png_data = {
        // PNG文件签名
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        
        // IHDR块 (图像头部信息)
        0x00, 0x00, 0x00, 0x0D,  // 块长度: 13字节
        0x49, 0x48, 0x44, 0x52,  // 块类型: "IHDR"
        0x00, 0x00, 0x00, 0x0A,  // 宽度: 10像素
        0x00, 0x00, 0x00, 0x0A,  // 高度: 10像素
        0x08,                    // 位深度: 8位
        0x02,                    // 颜色类型: 2 (RGB)
        0x00,                    // 压缩方法: 0
        0x00,                    // 滤波方法: 0
        0x00,                    // 隔行扫描: 0 (无隔行)
        0x7E, 0x9B, 0x55, 0x25,  // CRC
        
        // IDAT块 (图像数据)
        0x00, 0x00, 0x00, 0x1D,  // 块长度: 29字节
        0x49, 0x44, 0x41, 0x54,  // 块类型: "IDAT"
        // 压缩的RGB数据 (10x10红色像素)
        0x78, 0x9C, 0x62, 0xF8, 0xCF, 0xC0, 0xC0, 0xC0,
        0xC4, 0xC0, 0xC0, 0xC0, 0x00, 0xC6, 0x08, 0xC5,
        0x18, 0x18, 0x18, 0x01, 0x03, 0x03, 0x03, 0x00,
        0x00, 0xB4, 0x00, 0x1E,
        0x5B, 0x1C, 0x06, 0x71,  // CRC
        
        // IEND块 (文件结束)
        0x00, 0x00, 0x00, 0x00,  // 块长度: 0字节
        0x49, 0x45, 0x4E, 0x44,  // 块类型: "IEND"
        0xAE, 0x42, 0x60, 0x82   // CRC
    };
    return png_data;
}

int main() {
    std::cout << "FastExcel 图片插入修复测试" << std::endl;
    std::cout << "=========================" << std::endl;
    
    // 初始化FastExcel
    if (!fastexcel::initialize("logs/test_image_fix.log", true)) {
        std::cerr << "FastExcel初始化失败" << std::endl;
        return 1;
    }
    
    try {
        // 创建工作簿
        auto workbook = Workbook::create(Path("test_image_fix.xlsx"));
        if (!workbook) {
            std::cerr << "无法创建工作簿" << std::endl;
            return 1;
        }
        
        auto worksheet = workbook->addSheet("测试图片");
        
        // 添加标题
        worksheet->setValue(0, 0, std::string("图片插入测试"));
        worksheet->setValue(2, 0, std::string("测试图片:"));
        
        // 创建测试图片
        auto test_png_data = createTestPNG();
        std::cout << "创建测试PNG图片，大小: " << test_png_data.size() << " 字节" << std::endl;
        
        // 从内存数据创建图片对象
        auto image = Image::fromData(test_png_data, ImageFormat::PNG, "test.png");
        if (image) {
            image->setName("测试图片");
            image->setDescription("用于验证图片插入功能的测试图片");
            
            std::cout << "成功创建图片对象" << std::endl;
            std::cout << "图片格式: PNG" << std::endl;
            std::cout << "图片尺寸: " << image->getOriginalWidth() << "x" << image->getOriginalHeight() << std::endl;
            
            // 设置测试图片的显示大小为100x100像素（虽然原图是1x1）
            image->setCellAnchor(2, 1, 100, 100);  // 100x100像素显示大小
            
            // 插入图片到B3单元格
            std::string image_id = worksheet->insertImage(2, 1, std::move(image));
            if (!image_id.empty()) {
                std::cout << "成功插入图片到B3单元格，ID: " << image_id << std::endl;
            } else {
                std::cout << "图片插入失败" << std::endl;
            }
        } else {
            std::cout << "无法创建图片对象" << std::endl;
        }
        
        // 如果存在真实的图片文件，也测试一下
        std::string real_image_path = "tinaimage.png";
        std::ifstream test_file(real_image_path);
        if (test_file.good()) {
            test_file.close();
            
            worksheet->setValue(5, 0, std::string("真实图片:"));
            
            auto real_image = Image::fromFile(real_image_path);
            if (real_image) {
                real_image->setName("真实图片");
                real_image->setDescription("从文件加载的真实图片");
                
                std::cout << "\n从文件加载真实图片: " << real_image_path << std::endl;
                std::cout << "图片尺寸: " << real_image->getOriginalWidth() << "x" << real_image->getOriginalHeight() << std::endl;
                std::cout << "图片数据大小: " << real_image->getDataSize() << " 字节" << std::endl;
                
                // 插入到B6单元格，设置显示大小
                real_image->setCellAnchor(5, 1, 200, 150);  // 200x150像素显示大小
                std::string real_image_id = worksheet->insertImage(5, 1, std::move(real_image));
                if (!real_image_id.empty()) {
                    std::cout << "成功插入真实图片到B6单元格，ID: " << real_image_id << std::endl;
                } else {
                    std::cout << "真实图片插入失败" << std::endl;
                }
            }
        } else {
            std::cout << "\n未找到真实图片文件: " << real_image_path << std::endl;
        }
        
        // 保存工作簿
        std::cout << "\n正在保存工作簿..." << std::endl;
        if (workbook->save()) {
            std::cout << "工作簿保存成功: test_image_fix.xlsx" << std::endl;
            std::cout << "\n请使用Excel打开 test_image_fix.xlsx 文件验证图片是否正确显示" << std::endl;
        } else {
            std::cout << "工作簿保存失败" << std::endl;
        }
        
    } catch (const std::exception& e) {
        std::cerr << "测试过程中发生错误: " << e.what() << std::endl;
        fastexcel::cleanup();
        return 1;
    }
    
    // 清理资源
    fastexcel::cleanup();
    
    std::cout << "\n测试完成！" << std::endl;
    return 0;
}