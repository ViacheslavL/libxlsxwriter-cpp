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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hyperlink17.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    /* URL with whitespace. */
    worksheet->write_url(CELL("A1"), "http://google.com/some link", NULL);

    int result = workbook->close(); return result;
}
