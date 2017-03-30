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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hyperlink10.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format *format = workbook->add_format();

    format->set_underline(xlsxwriter::LXW_UNDERLINE_SINGLE);
    format->set_font_color(xlsxwriter::LXW_COLOR_RED);

    worksheet->write_url(CELL("A1"), "http://www.perl.org/", format);

    int result = workbook->close(); return result;
}
