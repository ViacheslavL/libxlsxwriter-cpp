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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink03.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();

    worksheet_write_url(worksheet1, CELL("A1"),  "http://www.perl.org/", NULL);
    worksheet_write_url(worksheet1, CELL("D4"),  "http://www.perl.org/", NULL);
    worksheet_write_url(worksheet1, CELL("A8"),  "http://www.perl.org/", NULL);
    worksheet_write_url(worksheet1, CELL("B6"),  "http://www.cpan.org/", NULL);
    worksheet_write_url(worksheet1, CELL("F12"), "http://www.cpan.org/", NULL);

    worksheet_write_url(worksheet2, CELL("C2"),  "http://www.google.com/", NULL);
    worksheet_write_url(worksheet2, CELL("C5"),  "http://www.cpan.org/",   NULL);
    worksheet_write_url(worksheet2, CELL("C7"),  "http://www.perl.org/",   NULL);

    int result = workbook->close(); return result;
}
