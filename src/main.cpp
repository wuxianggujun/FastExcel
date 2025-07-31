#include "Logger.hpp"
#include <iostream>

int main()
{
    // 初始化日志系统
    fastexcel::Logger::getInstance().initialize("logs/fastexcel.log", 
                                               fastexcel::Logger::Level::DEBUG, 
                                               true);
    
    // 测试不同级别的日志
    LOG_TRACE("这是一条跟踪日志");
    LOG_DEBUG("这是一条调试日志");
    LOG_INFO("这是一条信息日志");
    LOG_WARN("这是一条警告日志");
    LOG_ERROR("这是一条错误日志");
    LOG_CRITICAL("这是一条严重错误日志");
    
    // 测试带参数的日志
    std::string user = "张三";
    int age = 25;
    LOG_INFO("用户信息: 姓名={}, 年龄={}", user, age);
    
    // 测试中文字符
    LOG_INFO("FastExcel 日志系统初始化完成! 支持中文输出!");
    
    // 刷新日志
    fastexcel::Logger::getInstance().flush();
    
    std::cout << "日志测试完成，请查看控制台输出和日志文件！" << std::endl;
    
    // 显式关闭日志系统，避免析构顺序问题
    fastexcel::Logger::getInstance().shutdown();
    
    return 0;
}
