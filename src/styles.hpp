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

namespace xlsxwriter {

/*
 * Struct to represent a styles.
 */
struct styles : public xmlwriter{

public:
    void assemble_xml_file();

    /* Declarations required for unit testing. */
    #ifdef TESTING

    #endif /* TESTING */

    void _xml_declaration();
    void _write_style_sheet();
    void _write_font_size(uint16_t font_size);
    void _write_font_color_theme(uint8_t theme);
    void _write_font_name(const char *font_name);
    void _write_font_family(uint8_t font_family);
    void _write_font_scheme(const char *font_scheme);
    void _write_font(format *format);
    void _write_fonts();
    void _write_default_fill(const char *pattern);
    void _write_fills();
    void _write_border(format *format);
    void _write_borders();
    void _write_style_xf();
    void _write_cell_style_xfs();
    void _write_xf(format *format);
    void _write_cell_xfs();
    void _write_cell_style();
    void _write_cell_styles();
    void _write_dxfs();
    void _write_table_styles();

private:

    FILE *file;
    uint32_t font_count;
    uint32_t xf_count;
    uint32_t dxf_count;
    uint32_t num_format_count;
    uint32_t border_count;
    uint32_t fill_count;
    lxw_formats *xf_formats;
    lxw_formats *dxf_formats;

};



} // namespace xlsxwriter

#endif /* __LXW_STYLES_HPP__ */
