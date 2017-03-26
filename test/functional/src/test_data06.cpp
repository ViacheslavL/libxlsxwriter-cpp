/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case to test data writing.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_data06.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format    *format1   = workbook->add_format();
    xlsxwriter::format    *format2   = workbook->add_format();
    xlsxwriter::format    *format3   = workbook->add_format();

    format1->bold = 1;

    format2->italic = 1;

    format3->bold = 1;
    format3->italic = 1;

    worksheet->write_string(CELL("A1"), "Foo", format1);
    worksheet->write_string(CELL("A2"), "Bar", format2);
    worksheet->write_string(CELL("A3"), "Baz", format3);

    int result = workbook->close(); return result;
}

