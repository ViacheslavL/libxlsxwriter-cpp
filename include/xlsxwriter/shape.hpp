#ifndef __LXW_SHAPE_H__
#define __LXW_SHAPE_H__

#include "common.hpp"
#include <stdint.h>
#include "format.hpp"

namespace xlsxwriter {

enum dash_types {
    SOLID,
    ROUND_BOT,
    SQUARE_DOT,
    DASH,
    DASH_DOT,
    LONG_DASH,
    LONG_DASH_DOT,
    LONG_DASH_DOT_DOT,
    DOT,
    SYSTEM_DASH_DOT,
    SYSTEM_DASH_DOT_DOT
};

enum shape_pattern_types {
    SHAPE_PATTERN_PERCENT_5,
    SHAPE_PATTERN_PERCENT_10,
    SHAPE_PATTERN_PERCENT_20,
    SHAPE_PATTERN_PERCENT_25,
    SHAPE_PATTERN_PERCENT_30,
    SHAPE_PATTERN_PERCENT_40,
    SHAPE_PATTERN_PERCENT_50,
    SHAPE_PATTERN_PERCENT_60,
    SHAPE_PATTERN_PERCENT_70,
    SHAPE_PATTERN_PERCENT_75,
    SHAPE_PATTERN_PERCENT_80,
    SHAPE_PATTERN_PERCENT_90,
    SHAPE_PATTERN_LIGHT_DOWNWARD_DIAGONAL,
    SHAPE_PATTERN_LIGHT_UPWARD_DIAGONAL,
    SHAPE_PATTERN_DARK_DOWNWARD_DIAGONAL,
    SHAPE_PATTERN_DARK_UPWARD_DIAGONAL,
    SHAPE_PATTERN_WIDE_DOWNWARD_DIAGONAL,
    SHAPE_PATTERN_WIDE_UPWARD_DIAGONAL,
    SHAPE_PATTERN_LIGHT_VERTICAL,
    SHAPE_PATTERN_LIGHT_HORIZONTAL,
    SHAPE_PATTERN_NARROW_VERTICAL,
    SHAPE_PATTERN_NARROW_HORIZONTAL,
    SHAPE_PATTERN_DARK_VERTICAL,
    SHAPE_PATTERN_DARK_HORIZONTAL,
    SHAPE_PATTERN_DASHED_DOWNWARD_DIAGONAL,
    SHAPE_PATTERN_DASHED_UPWARD_DIAGONAL,
    SHAPE_PATTERN_DASHED_HORIZONTAL,
    SHAPE_PATTERN_DASHED_VERTICAL,
    SHAPE_PATTERN_SMALL_CONFETTI,
    SHAPE_PATTERN_LARGE_CONFETTI,
    SHAPE_PATTERN_ZIGZAG,
    SHAPE_PATTERN_WAVE,
    SHAPE_PATTERN_DIAGONAL_BRICK,
    SHAPE_PATTERN_HORIZONTAL_BRICK,
    SHAPE_PATTERN_WEAVE,
    SHAPE_PATTERN_PLAID,
    SHAPE_PATTERN_DIVOT,
    SHAPE_PATTERN_DOTTED_GRID,
    SHAPE_PATTERN_DOTTED_DIAMON,
    SHAPE_PATTERN_SHINGLE,
    SHAPE_PATTERN_TRELLIS,
    SHAPE_PATTERN_SPHERE,
    SHAPE_PATTERN_SMALL_GRID,
    SHAPE_PATTERN_LARGE_GRID,
    SHAPE_PATTERN_SMALL_CHECK,
    SHAPE_PATTERN_LARGE_CHECK,
    SHAPE_PATTERN_OUTLINED_DIAMOND,
    SHAPE_PATTERN_SOLID_DIAMON
};

struct lxw_pattern {
    lxw_pattern () : defined(false), pattern(0), fg_color(0), bg_color(0) {}
    bool defined;
    uint8_t pattern;
    lxw_color_t fg_color;
    lxw_color_t bg_color;
};

struct lxw_line {
    lxw_line() : defined(false), none(false), dash_type(0), color(0), transparency(0) {}
    bool defined;
    bool none;
    uint8_t dash_type;
    lxw_color_t color;
    double transparency;
};

struct lxw_shape_fill {
    lxw_shape_fill() : defined(false), none(false), color(0), transparency(0) {}
    bool defined;
    bool none;
    lxw_color_t color;
    double transparency;
};

struct lxw_shape_properties {
    lxw_shape_properties() {}
    lxw_shape_fill fill;
    lxw_line line;
    lxw_pattern pattern;
};

} // namespace xlsxwriter

#endif
