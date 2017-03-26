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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink15.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_write_url(worksheet, CELL("B2"), "external:subdir/blank.xlsx", NULL);

    int result = workbook->close(); return result;
}
