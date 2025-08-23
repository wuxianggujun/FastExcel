#include "UnifiedXMLGenerator.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace xml {

// ========== 工厂方法 ==========

std::unique_ptr<UnifiedXMLGenerator> UnifiedXMLGenerator::fromWorkbook(const core::Workbook* workbook) {
    GenerationContext context;
    context.workbook = workbook;
    // 防御式空指针检查：workbook 可能为空或内部组件尚未初始化
    if (workbook) {
        context.format_repo = &workbook->getStyles();
        context.sst = workbook->getSharedStrings(); // 获取SharedStringTable
    } else {
        context.format_repo = nullptr;
        context.sst = nullptr;
    }
    
    return std::make_unique<UnifiedXMLGenerator>(context);
}

std::unique_ptr<UnifiedXMLGenerator> UnifiedXMLGenerator::fromWorksheet(const core::Worksheet* worksheet) {
    GenerationContext context;
    context.worksheet = worksheet;
    if (worksheet) {
        context.workbook = worksheet->getParentWorkbook().get();
        if (context.workbook) {
            context.format_repo = &context.workbook->getStyles();
        }
    }
    
    return std::make_unique<UnifiedXMLGenerator>(context);
}


// ========== XMLGeneratorFactory 实现 ==========

std::unique_ptr<UnifiedXMLGenerator> XMLGeneratorFactory::createLightweightGenerator() {
    UnifiedXMLGenerator::GenerationContext context;
    // 轻量级生成器不需要完整的context
    return std::make_unique<UnifiedXMLGenerator>(context);
}

}} // namespace fastexcel::xml
