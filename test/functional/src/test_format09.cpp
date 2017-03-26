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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_format09.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *border1   = workbook->add_format();
    xlsxwriter::format    *border2   = workbook->add_format();
    xlsxwriter::format    *border3   = workbook->add_format();
    xlsxwriter::format    *border4   = workbook->add_format();


    format_set_border      (border1, LXW_BORDER_HAIR);
    format_set_border_color(border1, xlsxwriter::LXW_COLOR_RED);

    format_set_diag_type (border2, LXW_DIAGONAL_BORDER_UP);
    format_set_diag_color(border2, xlsxwriter::LXW_COLOR_RED);

    format_set_diag_type (border3, LXW_DIAGONAL_BORDER_DOWN);
    format_set_diag_color(border3, xlsxwriter::LXW_COLOR_RED);

    format_set_diag_type (border4, LXW_DIAGONAL_BORDER_UP_DOWN);
    format_set_diag_color(border4, xlsxwriter::LXW_COLOR_RED);

    worksheet_write_blank(worksheet, 1, 1, border1);
    worksheet_write_blank(worksheet, 3, 1, border2);
    worksheet_write_blank(worksheet, 5, 1, border3);
    worksheet_write_blank(worksheet, 7, 1, border4);

    int result = workbook->close(); return result;
}
