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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_set_start_page02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_set_start_page(worksheet, 2);
    worksheet_set_paper(worksheet, 9);
    worksheet->vertical_dpi = 200;

    worksheet->write_string(CELL("A1"), "Foo" , NULL);

    int result = workbook->close(); return result;
}
