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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink05.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_write_url(    worksheet, CELL("A1"), "http://www.perl.org/", NULL);
    worksheet_write_url_opt(worksheet, CELL("A3"), "http://www.perl.org/", NULL, "Perl home", NULL);
    worksheet_write_url_opt(worksheet, CELL("A5"), "http://www.perl.org/", NULL, "Perl home", "Tool Tip");
    worksheet_write_url_opt(worksheet, CELL("A7"), "http://www.cpan.org/", NULL, "CPAN",      "Download");

    int result = workbook->close(); return result;
}
