/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * styles - A libxlsxwriter library for creating Excel XLSX styles files.
 *
 */
#ifndef __LXW_STYLES_HPP__
#define __LXW_STYLES_HPP__

#include <stdint.h>

#include "format.hpp"
#include "xmlwriter.hpp"
#include <vector>

namespace xlsxwriter {

class packager;
/*
 * Struct to represent a styles.
 */
class styles : public xmlwriter{

    friend class packager;
public:
    void assemble_xml_file();

    /* Declarations required for unit testing. */
    #ifdef TESTING

    #endif /* TESTING */

    void _xml_declaration();
    void _write_style_sheet();
    void _write_font_size(uint16_t font_size);
    void _write_font_color_theme(uint8_t theme);
    void _write_font_name(const std::string& font_name);
    void _write_font_family(uint8_t font_family);
    void _write_font_scheme(const std::string& font_scheme);
    void _write_font(const format_ptr& format);
    void _write_fonts();
    void _write_default_fill(const std::string& pattern);
    void _write_fills();
    void _write_border(const format_ptr& format);
    void _write_borders();
    void _write_style_xf();
    void _write_cell_style_xfs();
    void _write_xf(const format_ptr& format);
    void _write_cell_xfs();
    void _write_cell_style();
    void _write_cell_styles();
    void _write_dxfs();
    void _write_table_styles();

private:
    uint32_t font_count;
    uint32_t xf_count;
    uint32_t dxf_count;
    uint32_t num_format_count;
    uint32_t border_count;
    uint32_t fill_count;
    std::vector<format_ptr> xf_formats;
    std::vector<format_ptr> dxf_formats;

    void _write_num_fmt(uint16_t num_fmt_id, char* format_code);
    void _write_num_fmts();
    void _write_font_color_rgb(int32_t rgb);
    uint8_t _has_alignment(const format_ptr &format);
    void _write_alignment(const format_ptr &format);
    void _write_vert_align(const std::string &align);
    void _write_font_underline(uint8_t underline);
    uint8_t _apply_alignment(const format_ptr &format);
    void _write_protection(const format_ptr &format);
    void _write_fg_color(lxw_color_t color);
    void _write_sub_border(const std::string &type, uint8_t style, lxw_color_t color);
    void _write_fill(const format_ptr &format);
    void _write_bg_color(lxw_color_t color);
    void _write_border_color(lxw_color_t color);
};



} // namespace xlsxwriter

#endif /* __LXW_STYLES_HPP__ */
