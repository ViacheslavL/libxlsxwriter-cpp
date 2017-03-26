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

    xlsxwriter::workbook *workbook   = new xlsxwriter::workbook("test_repeat05.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet3 = workbook->add_worksheet();

    worksheet1->set_paper(9);
    worksheet1->set_vertical_dpi(200);

    worksheet3->set_paper(9);
    worksheet3->set_vertical_dpi(200);

    (void) worksheet2;

    worksheet1->repeat_rows(0, 0);
    worksheet3->repeat_rows(2, 3);
    worksheet3->repeat_columns(1, 5);

    worksheet1->write_string(CELL("A1"), "Foo" );

    int result = workbook->close(); return result;
}
