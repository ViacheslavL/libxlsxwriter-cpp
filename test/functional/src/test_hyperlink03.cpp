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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hyperlink03.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();

    worksheet1->write_url(CELL("A1"),  "http://www.perl.org/", NULL);
    worksheet1->write_url(CELL("D4"),  "http://www.perl.org/", NULL);
    worksheet1->write_url(CELL("A8"),  "http://www.perl.org/", NULL);
    worksheet1->write_url(CELL("B6"),  "http://www.cpan.org/", NULL);
    worksheet1->write_url(CELL("F12"), "http://www.cpan.org/", NULL);

    worksheet2->write_url(CELL("C2"),  "http://www.google.com/", NULL);
    worksheet2->write_url(CELL("C5"),  "http://www.cpan.org/",   NULL);
    worksheet2->write_url(CELL("C7"),  "http://www.perl.org/",   NULL);

    int result = workbook->close(); return result;
}
