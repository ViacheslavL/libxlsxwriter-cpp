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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_simple02.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet("Data Sheet");
    xlsxwriter::worksheet *worksheet3 = workbook->add_worksheet();

    xlsxwriter::format    *format     = workbook->add_format();
    format->set_bold();

    worksheet1->write_string(0, 0, "Foo");
    worksheet1->write_number(1, 0, 123);

    worksheet3->write_string(1, 1, "Foo");
    worksheet3->write_string(2, 1, "Bar", format);
    worksheet3->write_number(3, 2, 234);

    (void)worksheet2; /* Unused. For testing only. */

    int result = workbook->close(); return result;
}
