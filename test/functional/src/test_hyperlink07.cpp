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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink07.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_write_url_opt(worksheet, CELL("A1"), "external:\\\\VBOXSVR\\share\\foo.xlsx", NULL, "J:\\foo.xlsx", NULL);
    worksheet_write_url(    worksheet, CELL("A3"), "external:foo.xlsx" , NULL);

    int result = workbook->close(); return result;
}
