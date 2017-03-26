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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image23.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_insert_image(worksheet, CELL("B2"), "images/black_72.jpg");
    worksheet_insert_image(worksheet, CELL("B8"), "images/black_96.jpg");
    worksheet_insert_image(worksheet, CELL("B13"), "images/black_150.jpg");
    worksheet_insert_image(worksheet, CELL("B17"), "images/black_300.jpg");

    int result = workbook->close(); return result;
}
