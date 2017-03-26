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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink11.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format *format = workbook->add_format();

    format_set_underline(format, LXW_UNDERLINE_SINGLE);
    format_set_font_color(format, LXW_COLOR_BLUE);

    worksheet_write_url(worksheet, CELL("A1"), "http://www.perl.org/", format);

    int result = workbook->close(); return result;
}
