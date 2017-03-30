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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_format10.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *border1   = workbook->add_format();
    xlsxwriter::format    *border2   = workbook->add_format();
    xlsxwriter::format    *border3   = workbook->add_format();


    border1->set_bg_color(xlsxwriter::LXW_COLOR_RED);

    border2->set_bg_color(xlsxwriter::LXW_COLOR_YELLOW);
    border2->set_pattern (xlsxwriter::LXW_PATTERN_DARK_VERTICAL);

    border3->set_bg_color(xlsxwriter::LXW_COLOR_YELLOW);
    border3->set_fg_color(xlsxwriter::LXW_COLOR_RED);
    border3->set_pattern (xlsxwriter::LXW_PATTERN_GRAY_0625);

    worksheet->write_blank(1, 1, border1);
    worksheet->write_blank(3, 1, border2);
    worksheet->write_blank(5, 1, border3);

    int result = workbook->close(); return result;
}
