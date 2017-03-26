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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_default_row03.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    uint8_t row;

    worksheet->set_default_row(24, true);

    worksheet->write_string(CELL("A1"), "Foo" , NULL);
    worksheet->write_string(CELL("A10"), "Bar" , NULL);

    for (row = 1; row <= 8; row++)
        worksheet->set_row(row, 24, NULL);

    int result = workbook->close(); return result;
}
