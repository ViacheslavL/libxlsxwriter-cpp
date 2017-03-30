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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hyperlink05.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_url(CELL("A1"), "http://www.perl.org/", NULL);
    worksheet->write_url_opt(CELL("A3"), "http://www.perl.org/", NULL, "Perl home", NULL);
    worksheet->write_url_opt(CELL("A5"), "http://www.perl.org/", NULL, "Perl home", "Tool Tip");
    worksheet->write_url_opt(CELL("A7"), "http://www.cpan.org/", NULL, "CPAN",      "Download");

    int result = workbook->close(); return result;
}
