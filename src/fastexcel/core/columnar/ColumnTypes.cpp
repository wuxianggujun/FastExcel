#include "ColumnTypes.hpp"
#include <algorithm>

namespace fastexcel {
namespace core {
namespace columnar {

// ValidityBitmap 实现
void ValidityBitmap::setBit(uint32_t row) {
    if (row >= max_row_) {
        max_row_ = row + 1;
    }
    
    uint32_t word_index = row / 64;
    uint32_t bit_index = row % 64;
    
    if (word_index >= bits_.size()) {
        bits_.resize(word_index + 1, 0);
    }
    
    bits_[word_index] |= (1ULL << bit_index);
}

bool ValidityBitmap::getBit(uint32_t row) const {
    if (row >= max_row_) {
        return false;
    }
    
    uint32_t word_index = row / 64;
    uint32_t bit_index = row % 64;
    
    if (word_index >= bits_.size()) {
        return false;
    }
    
    return (bits_[word_index] & (1ULL << bit_index)) != 0;
}

void ValidityBitmap::clear() {
    bits_.clear();
    max_row_ = 0;
}

size_t ValidityBitmap::getMemoryUsage() const {
    return bits_.capacity() * sizeof(uint64_t);
}

}}} // namespace fastexcel::core::columnar