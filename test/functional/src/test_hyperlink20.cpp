/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Test to compare output against Excel files.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink20.xlsx");

    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format *format1 = workbook->add_format();
    xlsxwriter::format *format2 = workbook->add_format();

    format_set_underline(format1, LXW_UNDERLINE_SINGLE);
    format_set_font_color(format1, LXW_COLOR_BLUE);

    format_set_underline(format2, LXW_UNDERLINE_SINGLE);
    format_set_font_color(format2, xlsxwriter::LXW_COLOR_RED);


    worksheet_write_url(worksheet, CELL("A1"), "http://www.python.org/1", format1);
    worksheet_write_url(worksheet, CELL("A2"), "http://www.python.org/2", format2);

    int result = workbook->close(); return result;
}
