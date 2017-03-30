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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_format08.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *border1   = workbook->add_format();
    xlsxwriter::format    *border2   = workbook->add_format();
    xlsxwriter::format    *border3   = workbook->add_format();
    xlsxwriter::format    *border4   = workbook->add_format();
    xlsxwriter::format    *border5   = workbook->add_format();


    border1->set_bottom(xlsxwriter::LXW_BORDER_THIN);
    border1->set_bottom_color(xlsxwriter::LXW_COLOR_RED);

    border2->set_top(xlsxwriter::LXW_BORDER_THIN);
    border2->set_top_color(xlsxwriter::LXW_COLOR_RED);

    border3->set_left(xlsxwriter::LXW_BORDER_THIN);
    border3->set_left_color(xlsxwriter::LXW_COLOR_RED);

    border4->set_right(xlsxwriter::LXW_BORDER_THIN);
    border4->set_right_color(xlsxwriter::LXW_COLOR_RED);

    border5->set_border(xlsxwriter::LXW_BORDER_THIN);
    border5->set_border_color(xlsxwriter::LXW_COLOR_RED);

    worksheet->write_blank(1, 1, border1);
    worksheet->write_blank(3, 1, border2);
    worksheet->write_blank(5, 1, border3);
    worksheet->write_blank(7, 1, border4);
    worksheet->write_blank(9, 1, border5);

    int result = workbook->close(); return result;
}
