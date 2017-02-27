/*****************************************************************************
 * format - A library for creating Excel XLSX format files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xmlwriter.hpp"
#include "format.hpp"
#include "utility.hpp"

namespace xlsxwriter {

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new format object.
 */
format::format()
{
    xf_format_indices = NULL;

    xf_index = LXW_PROPERTY_UNSET;
    dxf_index = LXW_PROPERTY_UNSET;

    num_format_index = 0;
    font_index = 0;
    has_font = false;
    has_dxf_font = false;
    font_size = 11;
    bold = false;
    italic = false;
    font_color = LXW_COLOR_UNSET;
    underline = false;
    font_strikeout = false;
    font_outline = false;
    font_shadow = false;
    font_script = false;
    font_family = LXW_DEFAULT_FONT_FAMILY;
    font_charset = false;
    font_condense = false;
    font_extend = false;
    theme = false;
    hyperlink = false;

    hidden = false;
    locked = true;

    text_h_align = LXW_ALIGN_NONE;
    text_wrap = false;
    text_v_align = LXW_ALIGN_NONE;
    text_justlast = false;
    rotation = 0;

    fg_color = LXW_COLOR_UNSET;
    bg_color = LXW_COLOR_UNSET;
    pattern = LXW_PATTERN_NONE;
    has_fill = false;
    has_dxf_fill = false;
    fill_index = 0;
    fill_count = 0;

    border_index = 0;
    has_border = false;
    has_dxf_border = false;
    border_count = 0;

    bottom = LXW_BORDER_NONE;
    left = LXW_BORDER_NONE;
    right = LXW_BORDER_NONE;
    top = LXW_BORDER_NONE;
    diag_border = LXW_BORDER_NONE;
    diag_type = LXW_BORDER_NONE;
    bottom_color = LXW_COLOR_UNSET;
    left_color = LXW_COLOR_UNSET;
    right_color = LXW_COLOR_UNSET;
    top_color = LXW_COLOR_UNSET;
    diag_color = LXW_COLOR_UNSET;

    indent = 0;
    shrink = false;
    merge_range = false;
    reading_order = 0;
    just_distrib = false;
    color_indexed = false;
    font_only = false;
}

/*
 * Check a user input color.
 */
STATIC lxw_color_t
_check_color(lxw_color_t color)
{
    if (color == LXW_COLOR_UNSET)
        return color;
    else
        return color & LXW_COLOR_MASK;
}

/*
 * Check a user input border.
 */
STATIC uint8_t
_check_border(uint8_t border)
{
    if (border >= LXW_BORDER_THIN && border <= LXW_BORDER_SLANT_DASH_DOT)
        return border;
    else
        return LXW_BORDER_NONE;
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Returns a format struct suitable for hashing as a lookup key. This is
 * mainly a memcpy with any pointer members set to NULL.
 */
format* format::_get_format_key()
{
    format *key = new format();

    memcpy(key, this, sizeof(format));

    /* Set pointer members to NULL since they aren't part of the comparison. */
    key->xf_format_indices = NULL;
    key->num_xf_formats = NULL;
    key->list_pointers.stqe_next = NULL;

    return key;

mem_error:
    return NULL;
}

/*
 * Returns a font struct suitable for hashing as a lookup key.
 */
lxw_font * format::get_font_key()
{
    lxw_font *key = new lxw_font{};

    key->font_name = font_name;
    key->font_size = font_size;
    key->bold = bold;
    key->italic = italic;
    key->font_color = font_color;
    key->underline = underline;
    key->font_strikeout = font_strikeout;
    key->font_outline = font_outline;
    key->font_shadow = font_shadow;
    key->font_script = font_script;
    key->font_family = font_family;
    key->font_charset = font_charset;
    key->font_condense = font_condense;
    key->font_extend = font_extend;

    return key;

mem_error:
    return NULL;
}

/*
 * Returns a border struct suitable for hashing as a lookup key.
 */
lxw_border * format::get_border_key()
{
    lxw_border *key = new lxw_border{};

    key->bottom = bottom;
    key->left = left;
    key->right = right;
    key->top = top;
    key->diag_border = diag_border;
    key->diag_type = diag_type;
    key->bottom_color = bottom_color;
    key->left_color = left_color;
    key->right_color = right_color;
    key->top_color = top_color;
    key->diag_color = diag_color;

    return key;

mem_error:
    return NULL;
}

/*
 * Returns a pattern fill struct suitable for hashing as a lookup key.
 */
lxw_fill * format::get_fill_key()
{
    lxw_fill *key = new lxw_fill{};

    key->fg_color = fg_color;
    key->bg_color = bg_color;
    key->pattern = pattern;

    return key;
}

/*
 * Returns the XF index number used by Excel to identify a format.
 */
int32_t format::get_xf_index()
{
    format *format_key;
    format *existing_format;
    lxw_hash_element *hash_element;
    lxw_hash_table *formats_hash_table = xf_format_indices;
    int32_t index;

    /* Note: The formats_hash_table/xf_format_indices contains the unique and
     * more importantly the *used* formats in the workbook.
     */

    /* Format already has an index number so return it. */
    if (xf_index != LXW_PROPERTY_UNSET) {
        return xf_index;
    }

    /* Otherwise, the format doesn't have an index number so we assign one.
     * First generate a unique key to identify the format in the hash table.
     */
    format_key = _get_format_key();

    /* Return the default format index if the key generation failed. */
    if (!format_key)
        return 0;

    /* Look up the format in the hash table. */
    hash_element =
        lxw_hash_key_exists(formats_hash_table, format_key,
                            sizeof(format));

    if (hash_element) {
        /* Format matches existing format with an index. */
        delete format_key;
        existing_format = (format*)hash_element->value;
        return existing_format->xf_index;
    }
    else {
        /* New format requiring an index. */
        index = formats_hash_table->unique_count;
        xf_index = index;
        lxw_insert_hash_element(formats_hash_table, format_key, this,
                                sizeof(format));
        return index;
    }
}

/*
 * Set the font_name property.
 */
void format::set_font_name(const std::string& name)
{
    font_name = name;
}

/*
 * Set the font_size property.
 */
void format::set_font_size(uint16_t size)
{

    if (size >= LXW_MIN_FONT_SIZE && size <= LXW_MAX_FONT_SIZE)
        font_size = size;
}

/*
 * Set the font_color property.
 */
void format::set_font_color(lxw_color_t color)
{
    font_color = _check_color(color);
}

/*
 * Set the bold property.
 */
void format::set_bold()
{
    bold = true;
}

/*
 * Set the italic property.
 */

void format::set_italic()
{
    italic = true;
}

/*
 * Set the underline property.
 */
void format::set_underline(uint8_t style)
{
    if (style >= LXW_UNDERLINE_SINGLE
        && style <= LXW_UNDERLINE_DOUBLE_ACCOUNTING)
        underline = style;
}

/*
 * Set the font_strikeout property.
 */
void format::set_font_strikeout()
{
    font_strikeout = true;
}

/*
 * Set the font_script property.
 */
void format::set_font_script(uint8_t style)
{
    if (style >= LXW_FONT_SUPERSCRIPT && style <= LXW_FONT_SUBSCRIPT)
        font_script = style;
}

/*
 * Set the font_outline property.
 */
void format::set_font_outline()
{
    font_outline = true;
}

/*
 * Set the font_shadow property.
 */
void format::set_font_shadow()
{
    font_shadow = true;
}

/*
 * Set the num_format property.
 */
void format::set_num_format(const std::string& format)
{
    num_format = format;
}

/*
 * Set the unlocked property.
 */
void format::set_unlocked()
{
    locked = false;
}

/*
 * Set the hidden property.
 */
void format::set_hidden()
{
    hidden = true;
}

/*
 * Set the align property.
 */
void format::set_align(uint8_t value)
{
    if (value >= LXW_ALIGN_LEFT && value <= LXW_ALIGN_DISTRIBUTED) {
        text_h_align = value;
    }

    if (value >= LXW_ALIGN_VERTICAL_TOP
        && value <= LXW_ALIGN_VERTICAL_DISTRIBUTED) {
        text_v_align = value;
    }
}

/*
 * Set the text_wrap property.
 */
void format::set_text_wrap()
{
    text_wrap = true;
}

/*
 * Set the rotation property.
 */
void format::set_rotation(int16_t angle)
{
    /* Convert user angle to Excel angle. */
    if (angle == 270) {
        rotation = 255;
    }
    else if (angle >= -90 || angle <= 90) {
        if (angle < 0)
            angle = -angle + 90;

        rotation = angle;
    }
    else {
        LXW_WARN("Rotation rotation outside range: -90 <= angle <= 90.");
        rotation = 0;
    }
}

/*
 * Set the indent property.
 */
void format::set_indent(uint8_t value)
{
    indent = value;
}

/*
 * Set the shrink property.
 */
void format::set_shrink()
{
    shrink = true;
}

/*
 * Set the text_justlast property.
 */
void format::set_text_justlast()
{
    text_justlast = true;
}

/*
 * Set the pattern property.
 */
void format::set_pattern(uint8_t value)
{
    pattern = value;
}

/*
 * Set the bg_color property.
 */
void format::set_bg_color(lxw_color_t color)
{
    bg_color = _check_color(color);
}

/*
 * Set the fg_color property.
 */
void format::set_fg_color(lxw_color_t color)
{
    fg_color = _check_color(color);
}

/*
 * Set the border property.
 */
void format::set_border(uint8_t style)
{
    style = _check_border(style);
    bottom = style;
    top = style;
    left = style;
    right = style;
}

/*
 * Set the border_color property.
 */
void format::set_border_color(lxw_color_t color)
{
    color = _check_color(color);
    bottom_color = color;
    top_color = color;
    left_color = color;
    right_color = color;
}

/*
 * Set the bottom property.
 */
void format::set_bottom(uint8_t style)
{
    bottom = _check_border(style);
}

/*
 * Set the bottom_color property.
 */
void format::set_bottom_color(lxw_color_t color)
{
    bottom_color = _check_color(color);
}

/*
 * Set the left property.
 */
void format::set_left(uint8_t style)
{
    left = _check_border(style);
}

/*
 * Set the left_color property.
 */
void format::set_left_color(lxw_color_t color)
{
    left_color = _check_color(color);
}

/*
 * Set the right property.
 */
void format::set_right(uint8_t style)
{
    right = _check_border(style);
}

/*
 * Set the right_color property.
 */
void format::set_right_color(lxw_color_t color)
{
    right_color = _check_color(color);
}

/*
 * Set the top property.
 */
void format::set_top(uint8_t style)
{
    top = _check_border(style);
}

/*
 * Set the top_color property.
 */
void format::set_top_color(lxw_color_t color)
{
    top_color = _check_color(color);
}

/*
 * Set the diag_type property.
 */
void format::set_diag_type(uint8_t type)
{
    if (type >= LXW_DIAGONAL_BORDER_UP && type <= LXW_DIAGONAL_BORDER_UP_DOWN)
        diag_type = type;
}

/*
 * Set the diag_color property.
 */
void format::set_diag_color(lxw_color_t color)
{
    diag_color = _check_color(color);
}

/*
 * Set the diag_border property.
 */
void format::set_diag_border(uint8_t style)
{
    diag_border = style;
}

/*
 * Set the num_format_index property.
 */
void format::set_num_format_index(uint8_t value)
{
    num_format_index = value;
}

/*
 * Set the valign property.
 */
void format::set_text_v_align(uint8_t value)
{
    text_v_align = value;
}

/*
 * Set the reading_order property.
 */
void format::set_reading_order(uint8_t value)
{
    reading_order = value;
}

/*
 * Set the font_family property.
 */
void format::set_font_family(uint8_t value)
{
    font_family = value;
}

/*
 * Set the font_charset property.
 */
void format::set_font_charset(uint8_t value)
{
    font_charset = value;
}

/*
 * Set the font_scheme property.
 */
void format::set_font_scheme(const std::string& value)
{
    font_scheme = value;
}

/*
 * Set the font_condense property.
 */
void format::set_font_condense()
{
    font_condense = true;
}

/*
 * Set the font_extend property.
 */
void format::set_font_extend()
{
    font_extend = true;
}

/*
 * Set the theme property.
 */
void format::set_theme(uint8_t value)
{
    theme = value;
}

} // namespace xlsxwriter
