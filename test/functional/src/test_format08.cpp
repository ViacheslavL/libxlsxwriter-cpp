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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_format08.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *border1   = workbook->add_format();
    xlsxwriter::format    *border2   = workbook->add_format();
    xlsxwriter::format    *border3   = workbook->add_format();
    xlsxwriter::format    *border4   = workbook->add_format();
    xlsxwriter::format    *border5   = workbook->add_format();


    format_set_bottom(border1, LXW_BORDER_THIN);
    format_set_bottom_color(border1, xlsxwriter::LXW_COLOR_RED);

    format_set_top(border2, LXW_BORDER_THIN);
    format_set_top_color(border2, xlsxwriter::LXW_COLOR_RED);

    format_set_left(border3, LXW_BORDER_THIN);
    format_set_left_color(border3, xlsxwriter::LXW_COLOR_RED);

    format_set_right(border4, LXW_BORDER_THIN);
    format_set_right_color(border4, xlsxwriter::LXW_COLOR_RED);

    format_set_border(border5, LXW_BORDER_THIN);
    format_set_border_color(border5, xlsxwriter::LXW_COLOR_RED);

    worksheet_write_blank(worksheet, 1, 1, border1);
    worksheet_write_blank(worksheet, 3, 1, border2);
    worksheet_write_blank(worksheet, 5, 1, border3);
    worksheet_write_blank(worksheet, 7, 1, border4);
    worksheet_write_blank(worksheet, 9, 1, border5);

    int result = workbook->close(); return result;
}
