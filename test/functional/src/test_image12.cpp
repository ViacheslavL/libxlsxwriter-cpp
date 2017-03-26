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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image12.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->set_row(1, 75, NULL);
    worksheet->set_column(2, 2, 32, NULL);

    worksheet_insert_image(worksheet, CELL("C2"), "images/logo.png");

    int result = workbook->close(); return result;
}
