#include "fastexcel/utils/ModuleLoggers.hpp"
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
    context.format_repo = &workbook->getStyleRepository();
    context.sst = workbook->getSharedStringTable(); // 获取SharedStringTable
    
    return std::make_unique<UnifiedXMLGenerator>(context);
}

std::unique_ptr<UnifiedXMLGenerator> UnifiedXMLGenerator::fromWorksheet(const core::Worksheet* worksheet) {
    GenerationContext context;
    context.worksheet = worksheet;
    if (worksheet) {
        context.workbook = worksheet->getParentWorkbook().get();
        if (context.workbook) {
            context.format_repo = &context.workbook->getStyleRepository();
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
