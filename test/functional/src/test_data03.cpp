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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_data03.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_number(0,       16383, 123, NULL);
    worksheet->write_number(1048575, 16383, 456, NULL);

    int result = workbook->close(); return result;
}
