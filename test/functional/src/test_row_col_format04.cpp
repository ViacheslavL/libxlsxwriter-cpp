/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case to test worksheet set_row() and set_column().
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_row_col_format04.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *italic      = workbook->add_format();
    italic->set_italic();

    worksheet->set_column(0, 0, 8.43, italic);

    worksheet->write_string(0, 0, "Foo", NULL);

    int result = workbook->close(); return result;
}
