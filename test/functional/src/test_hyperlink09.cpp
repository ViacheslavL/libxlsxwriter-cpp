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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink09.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_write_url(worksheet, CELL("A1"), "external:..\\foo.xlsx" , NULL);
    worksheet_write_url(worksheet, CELL("A3"), "external:..\\foo.xlsx#Sheet1!A1" , NULL);
    worksheet_write_url_opt(worksheet, CELL("A5"), "external:\\\\VBOXSVR\\share\\foo.xlsx#Sheet1!B2", NULL, "J:\\foo.xlsx#Sheet1!B2", NULL);

    int result = workbook->close(); return result;
}
