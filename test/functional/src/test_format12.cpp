/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case for TODO.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_format12.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format *top_left_bottom = workbook->add_format();
    top_left_bottom->set_bottom(xlsxwriter::LXW_BORDER_THIN);
    top_left_bottom->set_left(xlsxwriter::LXW_BORDER_THIN);
    top_left_bottom->set_top(xlsxwriter::LXW_BORDER_THIN);

    xlsxwriter::format *top_bottom = workbook->add_format();
    top_bottom->set_bottom(xlsxwriter::LXW_BORDER_THIN);
    top_bottom->set_top(xlsxwriter::LXW_BORDER_THIN);

    xlsxwriter::format *top_left = workbook->add_format();
    top_left->set_left(xlsxwriter::LXW_BORDER_THIN);
    top_left->set_top(xlsxwriter::LXW_BORDER_THIN);

    xlsxwriter::format *unused = workbook->add_format();
    unused->set_left(xlsxwriter::LXW_BORDER_THIN);

    worksheet->write_string(CELL("B2"), "test", top_left_bottom);
    worksheet->write_string(CELL("D2"), "test", top_left);
    worksheet->write_string(CELL("F2"), "test", top_bottom);

    int result = workbook->close(); return result;
}
