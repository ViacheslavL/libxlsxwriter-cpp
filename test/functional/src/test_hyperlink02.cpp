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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hyperlink02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_url(CELL("A1"),  "http://www.perl.org/", NULL);
    worksheet->write_url(CELL("D4"),  "http://www.perl.org/", NULL);
    worksheet->write_url(CELL("A8"),  "http://www.perl.org/", NULL);
    worksheet->write_url(CELL("B6"),  "http://www.cpan.org/", NULL);
    worksheet->write_url(CELL("F12"), "http://www.cpan.org/", NULL);

    int result = workbook->close(); return result;
}
