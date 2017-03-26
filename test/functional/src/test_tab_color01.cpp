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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_tab_color01.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_string(CELL("A1"), "Foo" , NULL);
    worksheet->set_tab_color(xlsxwriter::LXW_COLOR_RED);

    int result = workbook->close(); return result;
}
