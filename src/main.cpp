#include "fastexcel/utils/Logger.hpp"
#include <iostream>

int main()
{
    // Initialize logging system
    fastexcel::Logger::getInstance().initialize("logs/fastexcel.log", 
                                               fastexcel::Logger::Level::DEBUG, 
                                               true);
    
    // Test different log levels
    LOG_TRACE("This is a trace log");
    LOG_DEBUG("This is a debug log");
    LOG_INFO("This is an info log");
    LOG_WARN("This is a warning log");
    LOG_ERROR("This is an error log");
    LOG_CRITICAL("This is a critical log");
    
    // Test parameterized logging
    std::string user = "John";
    int age = 25;
    LOG_INFO("User info: name={}, age={}", user, age);
    
    // Test Chinese characters
    LOG_INFO("FastExcel logging system initialized! Supports Chinese output!");
    
    // Flush logs
    fastexcel::Logger::getInstance().flush();
    
    std::cout << "Log test completed, please check console output and log file!" << std::endl;
    
    // Explicitly shutdown logging system to avoid destructor order issues
    fastexcel::Logger::getInstance().shutdown();
    
    return 0;
}
