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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hyperlink20.xlsx");

    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format *format1 = workbook->add_format();
    xlsxwriter::format *format2 = workbook->add_format();

    format1->set_underline(xlsxwriter::LXW_UNDERLINE_SINGLE);
    format1->set_font_color(xlsxwriter::LXW_COLOR_BLUE);

    format2->set_underline(xlsxwriter::LXW_UNDERLINE_SINGLE);
    format2->set_font_color(xlsxwriter::LXW_COLOR_RED);


    worksheet->write_url(CELL("A1"), "http://www.python.org/1", format1);
    worksheet->write_url(CELL("A2"), "http://www.python.org/2", format2);

    int result = workbook->close(); return result;
}
