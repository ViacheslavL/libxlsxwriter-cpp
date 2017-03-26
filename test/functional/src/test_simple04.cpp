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

    lxw_datetime datetime1 = {0,    0,  0, 12, 0, 0};
    lxw_datetime datetime2 = {2013, 1, 27,  0, 0, 0};

    /* Use deprecated constructor for testing. */
    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_simple04.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *format1   = workbook->add_format();
    xlsxwriter::format    *format2   = workbook->add_format();
    format1->set_num_format_index(20);
    format2->set_num_format_index(14);

    worksheet->set_column(0, 0, 12, NULL);

    worksheet->write_datetime(0, 0, &datetime1, format1);
    worksheet->write_datetime(1, 0, &datetime2, format2);

    int result = workbook->close(); return result;
}
