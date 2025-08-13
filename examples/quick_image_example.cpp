/**
 * @file quick_image_example.cpp
 * @brief FastExcel图片插入快速入门示例
 * 
 * 这是一个简单的示例，展示如何在Excel中插入图片的基本用法。
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/Image.hpp"

#include <iostream>
#include <vector>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    std::cout << "FastExcel 图片插入快速示例" << std::endl;
    
    // 初始化FastExcel
    if (!fastexcel::initialize()) {
        std::cerr << "FastExcel初始化失败" << std::endl;
        return 1;
    }
    
    try {
        // 1. 创建工作簿
        auto workbook = Workbook::create(Path("quick_image_example.xlsx"));
        if (!workbook) {
            std::cerr << "无法创建工作簿" << std::endl;
            return 1;
        }
        
        // 2. 获取默认工作表
        auto worksheet = workbook->getActiveSheet();
        
        // 3. 添加一些文本内容
        worksheet->setValue(0, 0, std::string("FastExcel图片插入示例"));
        worksheet->setValue(2, 0, std::string("图片将插入到B3单元格:"));
        
        // 4. 使用真实的图片文件
        std::string image_path = "tinaimage.png";
        
        // 5. 从文件创建图片对象
        auto image = Image::fromFile(image_path);
        if (!image) {
            std::cerr << "无法加载图片文件: " << image_path << std::endl;
            std::cerr << "请确保图片文件存在" << std::endl;
            return 1;
        }
        
        // 6. 设置图片属性
        image->setName("ChatGPT示例图片");
        image->setDescription("这是一个真实的PNG图片文件");
        
        std::cout << "成功加载图片: " << image->getName() << std::endl;
        std::cout << "图片尺寸: " << image->getOriginalWidth() << "x" << image->getOriginalHeight() << std::endl;
        std::cout << "图片格式: " << static_cast<int>(image->getFormat()) << std::endl;
        std::cout << "文件大小: " << image->getDataSize() << " 字节" << std::endl;
        
        // 7. 插入图片到B3单元格（行2，列1）
        std::string image_id = worksheet->insertImage(2, 1, std::move(image));
        
        if (!image_id.empty()) {
            std::cout << "成功插入图片到B3单元格，ID: " << image_id << std::endl;
        } else {
            std::cout << "图片插入失败" << std::endl;
            return 1;
        }
        
        // 8. 也可以使用Excel地址格式插入第二个图片（范围锚定）
        auto image2 = Image::fromFile(image_path);
        if (image2) {
            image2->setName("范围锚定图片");
            image2->setDescription("锚定到D3:F8范围的图片");
            std::string image_id2 = worksheet->insertImage(2, 3, 7, 5, std::move(image2)); // D3:F8
            if (!image_id2.empty()) {
                std::cout << "成功插入范围锚定图片到D3:F8，ID: " << image_id2 << std::endl;
            }
        }
        
        // 9. 显示图片统计信息
        std::cout << "工作表中的图片数量: " << worksheet->getImageCount() << std::endl;
        
        // 10. 保存工作簿
        if (workbook->save()) {
            std::cout << "Excel文件保存成功: quick_image_example.xlsx" << std::endl;
        } else {
            std::cout << "Excel文件保存失败" << std::endl;
            return 1;
        }
        
    } catch (const std::exception& e) {
        std::cerr << "发生错误: " << e.what() << std::endl;
        fastexcel::cleanup();
        return 1;
    }
    
    // 清理资源
    fastexcel::cleanup();
    
    std::cout << "\n示例完成！请打开 quick_image_example.xlsx 查看结果。" << std::endl;
    std::cout << "注意：由于使用了简单的测试图片数据，图片可能很小。" << std::endl;
    std::cout << "在实际项目中，请使用 Image::fromFile() 方法加载真实的图片文件。" << std::endl;
    
    return 0;
}