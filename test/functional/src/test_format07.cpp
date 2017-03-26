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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_format07.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *format1    = workbook->add_format();
    xlsxwriter::format    *format2    = workbook->add_format();
    xlsxwriter::format    *format3    = workbook->add_format();


    format_set_num_format(format1, "0.000");
    format_set_num_format(format2, "0.00000");
    format_set_num_format(format3, "0.000000");


    worksheet->write_number(0, 0, 1.2222, NULL);
    worksheet->write_number(1, 0, 1.2222, format1);
    worksheet->write_number(2, 0, 1.2222, format2);
    worksheet->write_number(3, 0, 1.2222, format3);
    worksheet->write_number(4, 0, 1.2222, NULL);

    int result = workbook->close(); return result;
}
