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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_escapes07.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_write_url(worksheet, CELL("A1"), "http://example.com/!\"$%&'( )*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~", NULL);

    int result = workbook->close(); return result;
}
