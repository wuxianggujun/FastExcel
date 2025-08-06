#pragma once

#include <string>
#include <ctime>
#include <chrono>
#include <sstream>
#include <iomanip>

namespace fastexcel {
namespace utils {

/**
 * @brief 时间工具类 - 统一处理时间相关操作
 */
class TimeUtils {
public:
    /**
     * @brief 获取当前时间的 std::tm 结构
     * @return 当前本地时间
     */
    static std::tm getCurrentTime() {
        std::time_t now = std::time(nullptr);
        std::tm result;
        
#ifdef _WIN32
        localtime_s(&result, &now);
#else
        result = *std::localtime(&now);
#endif
        return result;
    }
    
    /**
     * @brief 获取当前UTC时间的 std::tm 结构
     * @return 当前UTC时间
     */
    static std::tm getCurrentUTCTime() {
        std::time_t now = std::time(nullptr);
        std::tm result;
        
#ifdef _WIN32
        gmtime_s(&result, &now);
#else
        result = *std::gmtime(&now);
#endif
        return result;
    }
    
    /**
     * @brief 格式化时间为ISO 8601格式 (YYYY-MM-DDTHH:MM:SSZ)
     * @param time 时间结构
     * @return ISO 8601格式的时间字符串
     */
    static std::string formatTimeISO8601(const std::tm& time) {
        std::ostringstream oss;
        oss << std::put_time(&time, "%Y-%m-%dT%H:%M:%SZ");
        return oss.str();
    }
    
    /**
     * @brief 格式化时间为自定义格式
     * @param time 时间结构
     * @param format 格式字符串（如 "%Y-%m-%d %H:%M:%S"）
     * @return 格式化后的时间字符串
     */
    static std::string formatTime(const std::tm& time, const std::string& format = "%Y-%m-%d %H:%M:%S") {
        std::ostringstream oss;
        oss << std::put_time(&time, format.c_str());
        return oss.str();
    }
    
    /**
     * @brief 将 std::tm 转换为 std::time_t
     * @param tm_time 时间结构
     * @return time_t 时间戳
     */
    static std::time_t tmToTimeT(const std::tm& tm_time) {
        std::tm temp = tm_time; // mktime 可能会修改输入
        return std::mktime(&temp);
    }
    
    /**
     * @brief 计算两个时间之间的天数差
     * @param start 开始时间
     * @param end 结束时间
     * @return 天数差（可能为负数）
     */
    static int daysBetween(const std::tm& start, const std::tm& end) {
        std::time_t start_time = tmToTimeT(start);
        std::time_t end_time = tmToTimeT(end);
        
        double seconds_diff = std::difftime(end_time, start_time);
        return static_cast<int>(seconds_diff / (24 * 60 * 60));
    }
    
    /**
     * @brief 将日期时间转换为Excel序列号
     * Excel使用1900年1月1日作为起始日期（序列号1）
     * @param datetime 日期时间
     * @return Excel序列号
     */
    static double toExcelSerialNumber(const std::tm& datetime) {
        // Excel使用1900年1月1日作为起始日期（序列号1）
        std::tm epoch = {};
        epoch.tm_year = 100; // 1900年
        epoch.tm_mon = 0;    // 1月
        epoch.tm_mday = 1;   // 1日
        
        std::time_t epoch_time = tmToTimeT(epoch);
        std::time_t input_time = tmToTimeT(datetime);
        
        // 计算天数差
        double days = std::difftime(input_time, epoch_time) / (24 * 60 * 60);
        
        // Excel的bug：将1900年当作闰年，所以需要调整
        if (days >= 60) { // 1900年3月1日之后
            days += 1;
        }
        
        return days + 1; // Excel序列号从1开始
    }
    
    /**
     * @brief 创建指定日期的 std::tm 结构
     * @param year 年份（如2024）
     * @param month 月份（1-12）
     * @param day 日期（1-31）
     * @param hour 小时（0-23，默认0）
     * @param minute 分钟（0-59，默认0）
     * @param second 秒（0-59，默认0）
     * @return std::tm 结构
     */
    static std::tm createTime(int year, int month, int day, 
                             int hour = 0, int minute = 0, int second = 0) {
        std::tm result = {};
        result.tm_year = year - 1900;  // tm_year 是从1900年开始的年数
        result.tm_mon = month - 1;     // tm_mon 是0-11
        result.tm_mday = day;
        result.tm_hour = hour;
        result.tm_min = minute;
        result.tm_sec = second;
        result.tm_isdst = -1;          // 让系统决定是否为夏令时
        
        // 标准化时间结构
        std::mktime(&result);
        return result;
    }
    
    /**
     * @brief 获取高精度时间戳（毫秒）
     * @return 毫秒时间戳
     */
    static int64_t getTimestampMs() {
        auto now = std::chrono::system_clock::now();
        auto duration = now.time_since_epoch();
        return std::chrono::duration_cast<std::chrono::milliseconds>(duration).count();
    }
    
    /**
     * @brief RAII性能计时器
     */
    class PerformanceTimer {
    private:
        std::chrono::high_resolution_clock::time_point start_;
        std::string name_;
        
    public:
        explicit PerformanceTimer(const std::string& name = "Timer")
            : start_(std::chrono::high_resolution_clock::now()), name_(name) {}
        
        ~PerformanceTimer() {
            auto end = std::chrono::high_resolution_clock::now();
            auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start_);
            // 这里可以输出到日志或其他地方
            // LOG_INFO("{} took {} ms", name_, duration.count());
        }
        
        /**
         * @brief 获取已经过的时间（毫秒）
         * @return 毫秒数
         */
        int64_t elapsedMs() const {
            auto now = std::chrono::high_resolution_clock::now();
            auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(now - start_);
            return duration.count();
        }
    };
};

}} // namespace fastexcel::utils
