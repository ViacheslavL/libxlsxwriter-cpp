/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Test case for writing data in optimization mode.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook_options options = {1, NULL};

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_optimize21.xlsx", &options);
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_string(CELL("A1"), "Foo",     NULL);
    worksheet->write_string(CELL("C3"), " Foo",    NULL);
    worksheet->write_string(CELL("E5"), "Foo ",    NULL);
    worksheet->write_string(CELL("A7"), "\tFoo\t", NULL);

    int result = workbook->close(); return result;
}
