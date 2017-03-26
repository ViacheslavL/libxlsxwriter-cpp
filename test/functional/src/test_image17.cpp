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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image17.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->set_row(1, 96, NULL);
    worksheet->set_column(2, 2, 18, NULL);

    worksheet_insert_image(worksheet, CELL("C2"), "images/issue32.png");

    int result = workbook->close(); return result;
}
