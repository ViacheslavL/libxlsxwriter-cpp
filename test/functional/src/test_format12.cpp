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
    format_set_bottom(top_left_bottom, LXW_BORDER_THIN);
    format_set_left(top_left_bottom, LXW_BORDER_THIN);
    format_set_top(top_left_bottom, LXW_BORDER_THIN);

    xlsxwriter::format *top_bottom = workbook->add_format();
    format_set_bottom(top_bottom, LXW_BORDER_THIN);
    format_set_top(top_bottom, LXW_BORDER_THIN);

    xlsxwriter::format *top_left = workbook->add_format();
    format_set_left(top_left, LXW_BORDER_THIN);
    format_set_top(top_left, LXW_BORDER_THIN);

    xlsxwriter::format *unused = workbook->add_format();
    format_set_left(unused, LXW_BORDER_THIN);

    worksheet->write_string(CELL("B2"), "test", top_left_bottom);
    worksheet->write_string(CELL("D2"), "test", top_left);
    worksheet->write_string(CELL("F2"), "test", top_bottom);

    int result = workbook->close(); return result;
}
