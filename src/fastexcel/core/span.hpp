#pragma once

#include <cstddef>

namespace fastexcel {
namespace core {

// C++17兼容的轻量级span类
template<typename T>
class span {
public:
    using element_type = T;
    using size_type = std::size_t;
    using iterator = T*;
    using const_iterator = const T*;
    
    constexpr span() noexcept : data_(nullptr), size_(0) {}
    constexpr span(T* data, size_type size) noexcept : data_(data), size_(size) {}
    
    constexpr iterator begin() const noexcept { return data_; }
    constexpr iterator end() const noexcept { return data_ + size_; }
    constexpr const_iterator cbegin() const noexcept { return data_; }
    constexpr const_iterator cend() const noexcept { return data_ + size_; }
    
    constexpr T* data() const noexcept { return data_; }
    constexpr size_type size() const noexcept { return size_; }
    constexpr bool empty() const noexcept { return size_ == 0; }
    
    constexpr T& operator[](size_type idx) const { return data_[idx]; }
    constexpr T& front() const { return data_[0]; }
    constexpr T& back() const { return data_[size_ - 1]; }
    
private:
    T* data_;
    size_type size_;
};

}} // namespace fastexcel::core