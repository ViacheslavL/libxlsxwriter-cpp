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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_escapes05.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet("Start");
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet("A & B");

    (void)worksheet2;

    worksheet1->write_url_opt(CELL("A1"), "internal:'A & B'!A1", NULL, "Jump to A & B");

    int result = workbook->close(); return result;
}
