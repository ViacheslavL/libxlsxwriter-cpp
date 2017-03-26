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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image04.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_insert_image(worksheet, CELL("E9"), "images/red.bmp");

    int result = workbook->close(); return result;
}
