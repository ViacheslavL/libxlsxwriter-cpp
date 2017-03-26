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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image25.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_insert_image(worksheet, CELL("B2"), "images/black_150.png");

    int result = workbook->close(); return result;
}
