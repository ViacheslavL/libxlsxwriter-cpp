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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_write_url(worksheet, CELL("A1"),  "http://www.perl.org/", NULL);
    worksheet_write_url(worksheet, CELL("D4"),  "http://www.perl.org/", NULL);
    worksheet_write_url(worksheet, CELL("A8"),  "http://www.perl.org/", NULL);
    worksheet_write_url(worksheet, CELL("B6"),  "http://www.cpan.org/", NULL);
    worksheet_write_url(worksheet, CELL("F12"), "http://www.cpan.org/", NULL);

    int result = workbook->close(); return result;
}
