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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image05.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_insert_image(worksheet, CELL("A1"), "images/blue.png");
    worksheet_insert_image(worksheet, CELL("B3"), "images/red.jpg");
    worksheet_insert_image(worksheet, CELL("D5"), "images/yellow.jpg");
    worksheet_insert_image(worksheet, CELL("F9"), "images/grey.png");

    int result = workbook->close(); return result;
}
