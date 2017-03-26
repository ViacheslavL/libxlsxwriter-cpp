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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_data02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    /* Tests for the row range. */
    worksheet->write_number(0,       0, 123, NULL);
    worksheet->write_number(1048575, 0, 456, NULL);

    /* These should be ignored. */
    worksheet->write_number(-1,      0, 123, NULL);
    worksheet->write_number(1048576, 0, 456, NULL);

    int result = workbook->close(); return result;
}
