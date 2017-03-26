/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case to test data writing.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_format10.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *border1   = workbook->add_format();
    xlsxwriter::format    *border2   = workbook->add_format();
    xlsxwriter::format    *border3   = workbook->add_format();


    format_set_bg_color(border1, xlsxwriter::LXW_COLOR_RED);

    format_set_bg_color(border2, LXW_COLOR_YELLOW);
    format_set_pattern (border2, LXW_PATTERN_DARK_VERTICAL);

    format_set_bg_color(border3, LXW_COLOR_YELLOW);
    format_set_fg_color(border3, xlsxwriter::LXW_COLOR_RED);
    format_set_pattern (border3, LXW_PATTERN_GRAY_0625);

    worksheet_write_blank(worksheet, 1, 1, border1);
    worksheet_write_blank(worksheet, 3, 1, border2);
    worksheet_write_blank(worksheet, 5, 1, border3);

    int result = workbook->close(); return result;
}
