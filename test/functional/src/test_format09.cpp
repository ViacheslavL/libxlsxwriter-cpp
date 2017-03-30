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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_format09.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *border1   = workbook->add_format();
    xlsxwriter::format    *border2   = workbook->add_format();
    xlsxwriter::format    *border3   = workbook->add_format();
    xlsxwriter::format    *border4   = workbook->add_format();


    border1->set_border      (xlsxwriter::LXW_BORDER_HAIR);
    border1->set_border_color(xlsxwriter::LXW_COLOR_RED);

    border2->set_diag_type (xlsxwriter::LXW_DIAGONAL_BORDER_UP);
    border2->set_diag_color(xlsxwriter::LXW_COLOR_RED);

    border3->set_diag_type (xlsxwriter::LXW_DIAGONAL_BORDER_DOWN);
    border3->set_diag_color(xlsxwriter::LXW_COLOR_RED);

    border4->set_diag_type (xlsxwriter::LXW_DIAGONAL_BORDER_UP_DOWN);
    border4->set_diag_color(xlsxwriter::LXW_COLOR_RED);

    worksheet->write_blank(1, 1, border1);
    worksheet->write_blank(3, 1, border2);
    worksheet->write_blank(5, 1, border3);
    worksheet->write_blank(7, 1, border4);

    int result = workbook->close(); return result;
}
