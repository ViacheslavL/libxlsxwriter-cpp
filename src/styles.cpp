/*****************************************************************************
 * styles - A library for creating Excel XLSX styles files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <xlsxwriter/xmlwriter.hpp>
#include <xlsxwriter/styles.hpp>
#include <xlsxwriter/utility.hpp>

/*
 * Forward declarations.
 */

namespace xlsxwriter {


/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
void styles::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <styleSheet> element.
 */
void styles::_write_style_sheet()
{
    xml_attribute_list attributes = {
        {"xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    };

    lxw_xml_start_tag("styleSheet", attributes);
}

/*
 * Write the <numFmt> element.
 */
void styles::_write_num_fmt(uint16_t num_fmt_id, std::string& format_code)
{
    xml_attribute_list attributes = {
        {"numFmtId", std::to_string(num_fmt_id)},
        {"formatCode", format_code}
    };

    lxw_xml_empty_tag("numFmt", attributes);
}

/*
 * Write the <numFmts> element.
 */
void styles::_write_num_fmts()
{
    if (num_format_count == 0)
        return;

    xml_attribute_list attributes = {
        {"count", std::to_string(num_format_count)}
    };

    lxw_xml_start_tag("numFmts", attributes);

    /* Write the numFmts elements. */
    for (const auto& format : xf_formats) {

        /* Ignore built-in number formats, i.e., < 164. */
        if (format->num_format_index < 164)
            continue;

        _write_num_fmt(format->num_format_index, format->num_format);
    }

    lxw_xml_end_tag("numFmts");
}

/*
 * Write the <sz> element.
 */
void styles::_write_font_size(uint16_t font_size)
{
    xml_attribute_list attributes = {
        {"val", std::to_string(font_size)}
    };

    lxw_xml_empty_tag("sz", attributes);
}

/*
 * Write the <color> element for themes.
 */
void styles::_write_font_color_theme(uint8_t theme)
{
    xml_attribute_list attributes = {
        {"theme", std::to_string(theme)}
    };

    lxw_xml_empty_tag("color", attributes);
}

/*
 * Write the <color> element for RGB colors.
 */
void styles::_write_font_color_rgb(int32_t rgb)
{
    char rgb_str[LXW_ATTR_32];

    lxw_snprintf(rgb_str, LXW_ATTR_32, "FF%06X", rgb & LXW_COLOR_MASK);

    xml_attribute_list attributes = {
        {"rgb", rgb_str}
    };

    lxw_xml_empty_tag("color", attributes);
}

/*
 * Write the <name> element.
 */
void styles::_write_font_name(const std::string& font_name)
{
    xml_attribute_list attributes = {
        {"val", (font_name.empty() ? LXW_DEFAULT_FONT_NAME : font_name)}
    };

    lxw_xml_empty_tag("name", attributes);

}

/*
 * Write the <family> element.
 */
void styles::_write_font_family(uint8_t font_family)
{
    xml_attribute_list attributes = {
        {"val", std::to_string(font_family)}
    };

    lxw_xml_empty_tag("family", attributes);
}

/*
 * Write the <scheme> element.
 */
void styles::_write_font_scheme(const std::string& font_scheme)
{
    xml_attribute_list attributes = {
        {"val", font_scheme.empty() ? "minor" : font_scheme}
    };

    lxw_xml_empty_tag("scheme", attributes);
}

/*
 * Write the underline font element.
 */
void styles::_write_font_underline(uint8_t underline)
{
    xml_attribute_list attributes;

    /* Handle the underline variants. */
    if (underline == LXW_UNDERLINE_DOUBLE)
        attributes.push_back({"val", "double"});
    else if (underline == LXW_UNDERLINE_SINGLE_ACCOUNTING)
        attributes.push_back({"val", "singleAccounting"});
    else if (underline == LXW_UNDERLINE_DOUBLE_ACCOUNTING)
        attributes.push_back({"val", "doubleAccounting"});
    /* Default to single underline. */

    lxw_xml_empty_tag("u", attributes);
}

/*
 * Write the <vertAlign> font sub-element.
 */
void styles::_write_vert_align(const std::string& align)
{
    xml_attribute_list attributes = {
        {"val", align}
    };

    lxw_xml_empty_tag("vertAlign", attributes);
}

/*
 * Write the <font> element.
 */
void styles::_write_font(const format_ptr& format)
{
    lxw_xml_start_tag("font");

    if (format->bold)
        lxw_xml_empty_tag("b");

    if (format->italic)
        lxw_xml_empty_tag("i");

    if (format->font_strikeout)
        lxw_xml_empty_tag("strike");

    if (format->font_outline)
        lxw_xml_empty_tag("outline");

    if (format->font_shadow)
        lxw_xml_empty_tag("shadow");

    if (format->underline)
        _write_font_underline(format->underline);

    if (format->font_script == LXW_FONT_SUPERSCRIPT)
        _write_vert_align("superscript");

    if (format->font_script == LXW_FONT_SUBSCRIPT)
        _write_vert_align("subscript");

    if (format->font_size)
        _write_font_size(format->font_size);

    if (format->theme)
        _write_font_color_theme(format->theme);
    else if (format->font_color != LXW_COLOR_UNSET)
        _write_font_color_rgb(format->font_color);
    else
        _write_font_color_theme(LXW_DEFAULT_FONT_THEME);

    _write_font_name(format->font_name);
    _write_font_family(format->font_family);

    /* Only write the scheme element for the default font type if it
     * is a hyperlink. */
    if ((format->font_name.empty()
         || LXW_DEFAULT_FONT_NAME == format->font_name)
        && !format->hyperlink) {
        _write_font_scheme(format->font_scheme);
    }

    lxw_xml_end_tag("font");
}

/*
 * Write the <fonts> element.
 */
void styles::_write_fonts()
{
    xml_attribute_list attributes = {
        {"count", std::to_string(font_count)}
    };

    lxw_xml_start_tag("fonts", attributes);

    for (const auto& format : xf_formats) {
        if (format->has_font)
            _write_font(format);
    }

    lxw_xml_end_tag("fonts");
}

/*
 * Write the default <fill> element.
 */
void styles::_write_default_fill(const std::string& pattern)
{
    xml_attribute_list attributes = {
        {"patternType", pattern}
    };

    lxw_xml_start_tag("fill");
    lxw_xml_empty_tag("patternFill", attributes);
    lxw_xml_end_tag("fill");
}

/*
 * Write the <fgColor> element.
 */
void styles::_write_fg_color(lxw_color_t color)
{

    char rgb_str[LXW_ATTR_32];
    lxw_snprintf(rgb_str, LXW_ATTR_32, "FF%06X", color & LXW_COLOR_MASK);
    xml_attribute_list attributes = {
        {"rgb", rgb_str}
    };

    lxw_xml_empty_tag("fgColor", attributes);
}

/*
 * Write the <bgColor> element.
 */
void styles::_write_bg_color(lxw_color_t color)
{
    xml_attribute_list attributes;
    char rgb_str[LXW_ATTR_32];

    if (color == LXW_COLOR_UNSET) {
        attributes.push_back({"indexed", "64"});
    }
    else {
        lxw_snprintf(rgb_str, LXW_ATTR_32, "FF%06X", color & LXW_COLOR_MASK);
        attributes.push_back({"rgb", rgb_str});
    }

    lxw_xml_empty_tag("bgColor", attributes);
}

/*
 * Write the <fill> element.
 */
void styles::_write_fill(const format_ptr& format)
{
    xml_attribute_list attributes;

    uint8_t pattern = format->pattern;
    lxw_color_t bg_color = format->bg_color;
    lxw_color_t fg_color = format->fg_color;

    static std::vector<std::string> patterns = {
        "none",
        "solid",
        "mediumGray",
        "darkGray",
        "lightGray",
        "darkHorizontal",
        "darkVertical",
        "darkDown",
        "darkUp",
        "darkGrid",
        "darkTrellis",
        "lightHorizontal",
        "lightVertical",
        "lightDown",
        "lightUp",
        "lightGrid",
        "lightTrellis",
        "gray125",
        "gray0625",
    };

    lxw_xml_start_tag("fill");

    if (pattern)
        attributes.push_back({"patternType", patterns[pattern]});

    lxw_xml_start_tag("patternFill", attributes);

    if (fg_color != LXW_COLOR_UNSET)
        _write_fg_color(fg_color);

    _write_bg_color(bg_color);

    lxw_xml_end_tag("patternFill");
    lxw_xml_end_tag("fill");
}

/*
 * Write the <fills> element.
 */
void styles::_write_fills()
{
    xml_attribute_list attributes = {
        {"count", std::to_string(fill_count)}
    };

    lxw_xml_start_tag("fills", attributes);

    /* Write the default fills. */
    _write_default_fill("none");
    _write_default_fill("gray125");

    for (const auto& format : xf_formats) {
        if (format->has_fill)
            _write_fill(format);
    }

    lxw_xml_end_tag("fills");
}

/*
 * Write the border <color> element.
 */
void styles::_write_border_color(lxw_color_t color)
{
    xml_attribute_list attributes;
    char rgb_str[LXW_ATTR_32];


    if (color != LXW_COLOR_UNSET) {
        lxw_snprintf(rgb_str, LXW_ATTR_32, "FF%06X", color & LXW_COLOR_MASK);
        attributes.push_back({"rgb", rgb_str});
    }
    else {
        attributes.push_back({"auto", "1"});
    }

    lxw_xml_empty_tag("color", attributes);
}

/*
 * Write the <border> sub elements such as <right>, <top>, etc.
 */
void styles::_write_sub_border(const std::string& type, uint8_t style, lxw_color_t color)
{
    static const std::vector<std::string> border_styles = {
        "none",
        "thin",
        "medium",
        "dashed",
        "dotted",
        "thick",
        "double",
        "hair",
        "mediumDashed",
        "dashDot",
        "mediumDashDot",
        "dashDotDot",
        "mediumDashDotDot",
        "slantDashDot",
    };

    if (!style) {
        lxw_xml_empty_tag(type);
        return;
    }

    xml_attribute_list attributes = {
        {"style", border_styles[style]}
    };

    lxw_xml_start_tag(type, attributes);

    _write_border_color(color);

    lxw_xml_end_tag(type);
}

/*
 * Write the <border> element.
 */
void styles::_write_border(const format_ptr& format)
{
    xml_attribute_list attributes;

    /* Add attributes for diagonal borders. */
    if (format->diag_type == LXW_DIAGONAL_BORDER_UP) {
        attributes.push_back({"diagonalUp", "1"});
    }
    else if (format->diag_type == LXW_DIAGONAL_BORDER_DOWN) {
        attributes.push_back({"diagonalDown", "1"});
    }
    else if (format->diag_type == LXW_DIAGONAL_BORDER_UP_DOWN) {
        attributes.push_back({"diagonalUp", "1"});
        attributes.push_back({"diagonalDown", "1"});
    }

    /* Ensure that a default diag border is set if the diag type is set. */
    if (format->diag_type && !format->diag_border) {
        format->diag_border = 1;
    }

    /* Write the start border tag. */
    lxw_xml_start_tag("border", attributes);

    /* Write the <border> sub elements. */
    _write_sub_border("left", format->left, format->left_color);
    _write_sub_border("right", format->right, format->right_color);
    _write_sub_border("top", format->top, format->top_color);
    _write_sub_border("bottom", format->bottom, format->bottom_color);
    _write_sub_border("diagonal", format->diag_border, format->diag_color);

    lxw_xml_end_tag("border");
}

/*
 * Write the <borders> element.
 */
void styles::_write_borders()
{
    xml_attribute_list attributes = {
        {"count", std::to_string(border_count)}
    };

    lxw_xml_start_tag("borders", attributes);

    for (const auto& format : xf_formats) {
        if (format->has_border)
            _write_border(format);
    }

    lxw_xml_end_tag("borders");
}

/*
 * Write the <xf> element for styles.
 */
void styles::_write_style_xf()
{
    xml_attribute_list attributes = {
        {"numFmtId", "0"},
        {"fontId", "0"},
        {"fillId", "0"},
        {"borderId", "0"}
    };

    lxw_xml_empty_tag("xf", attributes);
}

/*
 * Write the <cellStyleXfs> element.
 */
void styles::_write_cell_style_xfs()
{
    xml_attribute_list attributes = {
        {"count", "1"}
    };

    lxw_xml_start_tag("cellStyleXfs", attributes);
    _write_style_xf();
    lxw_xml_end_tag("cellStyleXfs");
}

/*
 * Check if a format struct has alignment properties set and the
 * "applyAlignment" attribute should be set.
 */
uint8_t styles::_apply_alignment(const format_ptr& format)
{
    return format->text_h_align != LXW_ALIGN_NONE
        || format->text_v_align != LXW_ALIGN_NONE
        || format->indent != 0
        || format->rotation != 0
        || format->text_wrap != 0
        || format->shrink != 0 || format->reading_order != 0;
}

/*
 * Check if a format struct has alignment properties set apart from the
 * LXW_ALIGN_VERTICAL_BOTTOM which Excel treats as a default.
 */
uint8_t styles::_has_alignment(const format_ptr& format)
{
    return format->text_h_align != LXW_ALIGN_NONE
        || !(format->text_v_align == LXW_ALIGN_NONE ||
             format->text_v_align == LXW_ALIGN_VERTICAL_BOTTOM)
        || format->indent != 0
        || format->rotation != 0
        || format->text_wrap != 0
        || format->shrink != 0 || format->reading_order != 0;
}

/*
 * Write the <alignment> element.
 */
void styles::_write_alignment(const format_ptr& format)
{
    xml_attribute_list attributes;
    int16_t rotation = format->rotation;

    /* Indent is only allowed for horizontal left, right and distributed. */
    /* If it is defined for any other alignment or no alignment has been  */
    /* set then default to left alignment. */
    if (format->indent
        && format->text_h_align != LXW_ALIGN_LEFT
        && format->text_h_align != LXW_ALIGN_RIGHT
        && format->text_h_align != LXW_ALIGN_DISTRIBUTED) {
        format->text_h_align = LXW_ALIGN_LEFT;
    }

    /* Check for properties that are mutually exclusive. */
    if (format->text_wrap)
        format->shrink = 0;

    if (format->text_h_align == LXW_ALIGN_FILL)
        format->shrink = 0;

    if (format->text_h_align == LXW_ALIGN_JUSTIFY)
        format->shrink = 0;

    if (format->text_h_align == LXW_ALIGN_DISTRIBUTED)
        format->shrink = 0;

    if (format->text_h_align != LXW_ALIGN_DISTRIBUTED)
        format->just_distrib = 0;

    if (format->indent)
        format->just_distrib = 0;

    if (format->text_h_align == LXW_ALIGN_LEFT)
        attributes.push_back({"horizontal", "left"});

    if (format->text_h_align == LXW_ALIGN_CENTER)
        attributes.push_back({"horizontal", "center"});

    if (format->text_h_align == LXW_ALIGN_RIGHT)
        attributes.push_back({"horizontal", "right"});

    if (format->text_h_align == LXW_ALIGN_FILL)
        attributes.push_back({"horizontal", "fill"});

    if (format->text_h_align == LXW_ALIGN_JUSTIFY)
        attributes.push_back({"horizontal", "justify"});

    if (format->text_h_align == LXW_ALIGN_CENTER_ACROSS)
        attributes.push_back({"horizontal", "centerContinuous"});

    if (format->text_h_align == LXW_ALIGN_DISTRIBUTED)
        attributes.push_back({"horizontal", "distributed"});

    if (format->just_distrib)
        attributes.push_back({"justifyLastLine", "1"});

    if (format->text_v_align == LXW_ALIGN_VERTICAL_TOP)
        attributes.push_back({"vertical", "top"});

    if (format->text_v_align == LXW_ALIGN_VERTICAL_CENTER)
        attributes.push_back({"vertical", "center"});

    if (format->text_v_align == LXW_ALIGN_VERTICAL_JUSTIFY)
        attributes.push_back({"vertical", "justify"});

    if (format->text_v_align == LXW_ALIGN_VERTICAL_DISTRIBUTED)
        attributes.push_back({"vertical", "distributed"});

    if (format->indent)
        attributes.push_back({"indent", std::to_string(format->indent)});

    /* Map rotation to Excel values. */
    if (rotation) {
        if (rotation == 270)
            rotation = 255;
        else if (rotation < 0)
            rotation = -rotation + 90;

        attributes.push_back({"textRotation", std::to_string(rotation)});
    }

    if (format->text_wrap)
        attributes.push_back({"wrapText", "1"});

    if (format->shrink)
        attributes.push_back({"shrinkToFit", "1"});

    if (format->reading_order == 1)
        attributes.push_back({"readingOrder", "1"});

    if (format->reading_order == 2)
        attributes.push_back({"readingOrder", "2"});

    if (attributes.empty())
        lxw_xml_empty_tag("alignment", attributes);
}

/*
 * Write the <protection> element.
 */
void styles::_write_protection(const format_ptr& format)
{
    xml_attribute_list attributes;

    if (!format->locked)
        attributes.push_back({"locked", "0"});

    if (format->hidden)
        attributes.push_back({"hidden", "1"});

    lxw_xml_empty_tag("protection", attributes);
}

/*
 * Write the <xf> element.
 */
void styles::_write_xf(const format_ptr& format)
{

    uint8_t has_protection = (!format->locked) | format->hidden;
    uint8_t has_alignment = _has_alignment(format);
    uint8_t apply_alignment = _apply_alignment(format);

    xml_attribute_list attributes = {
        {"numFmtId", std::to_string(format->num_format_index)},
        {"fontId", std::to_string(format->font_index)},
        {"fillId", std::to_string(format->fill_index)},
        {"borderId", std::to_string(format->border_index)},
        {"xfId", "0"}
    };

    if (format->num_format_index > 0)
        attributes.push_back({"applyNumberFormat", "1"});

    /* Add applyFont attribute if XF format uses a font element. */
    if (format->font_index > 0)
        attributes.push_back({"applyFont", "1"});

    /* Add applyFill attribute if XF format uses a fill element. */
    if (format->fill_index > 0)
        attributes.push_back({"applyFill", "1"});

    /* Add applyBorder attribute if XF format uses a border element. */
    if (format->border_index > 0)
        attributes.push_back({"applyBorder", "1"});

    /* We can also have applyAlignment without a sub-element. */
    if (apply_alignment)
        attributes.push_back({"applyAlignment", "1"});

    if (has_protection)
        attributes.push_back({"applyProtection", "1"});

    /* Write XF with sub-elements if required. */
    if (has_alignment || has_protection) {
        lxw_xml_start_tag("xf", attributes);

        if (has_alignment)
            _write_alignment(format);

        if (has_protection)
            _write_protection(format);

        lxw_xml_end_tag("xf");
    }
    else {
        lxw_xml_empty_tag("xf", attributes);
    }
}

/*
 * Write the <cellXfs> element.
 */
void styles::_write_cell_xfs()
{
    xml_attribute_list attributes {
        {"count", std::to_string(xf_count)}
    };

    lxw_xml_start_tag("cellXfs", attributes);

    for (const auto& format : xf_formats) {
        _write_xf(format);
    }

    lxw_xml_end_tag("cellXfs");
}

/*
 * Write the <cellStyle> element.
 */
void styles::_write_cell_style()
{
    xml_attribute_list attributes = {
        { "name", "Normal"},
        {"xfId", "0"},
        {"builtinId", "0"}
    };

    lxw_xml_empty_tag("cellStyle", attributes);
}

/*
 * Write the <cellStyles> element.
 */
void styles::_write_cell_styles()
{
    xml_attribute_list attributes = {
        {"count", "1"}
    };

    lxw_xml_start_tag("cellStyles", attributes);
    _write_cell_style();
    lxw_xml_end_tag("cellStyles");
}

/*
 * Write the <dxfs> element.
 */
void styles::_write_dxfs()
{
    xml_attribute_list attributes = {
        {"count", "0"}
    };

    lxw_xml_empty_tag("dxfs", attributes);
}

/*
 * Write the <tableStyles> element.
 */
void styles::_write_table_styles()
{
    xml_attribute_list attributes = {
        {"count", "0"},
        {"defaultTableStyle", "TableStyleMedium9"},
        {"defaultPivotStyle", "PivotStyleLight16"}
    };

    lxw_xml_empty_tag("tableStyles", attributes);
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Assemble and write the XML file.
 */
void styles::assemble_xml_file()
{
    /* Write the XML declaration. */
    _xml_declaration();

    /* Add the style sheet. */
    _write_style_sheet();

    /* Write the number formats. */
    _write_num_fmts();

    /* Write the fonts. */
    _write_fonts();

    /* Write the fills. */
    _write_fills();

    /* Write the borders element. */
    _write_borders();

    /* Write the cellStyleXfs element. */
    _write_cell_style_xfs();

    /* Write the cellXfs element. */
    _write_cell_xfs();

    /* Write the cellStyles element. */
    _write_cell_styles();

    /* Write the dxfs element. */
    _write_dxfs();

    /* Write the tableStyles element. */
    _write_table_styles();

    /* Write the colors element. */
    /* _write_colors(self); */

    /* Close the style sheet tag. */
    lxw_xml_end_tag("styleSheet");
}


} // namespace xlsxwriter
/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
