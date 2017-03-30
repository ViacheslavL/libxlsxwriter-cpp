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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hide01.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet3 = workbook->add_worksheet();

    worksheet2->hide();

    /* Avoid warnings about unused variables. */
    (void)worksheet1;
    (void)worksheet3;

    int result = workbook->close(); return result;
}
