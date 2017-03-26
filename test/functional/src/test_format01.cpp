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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_format01.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet("Data Sheet");
    xlsxwriter::worksheet *worksheet3 = workbook->add_worksheet();

    xlsxwriter::format    *unused1    = workbook->add_format();
    xlsxwriter::format    *format     = workbook->add_format();
    xlsxwriter::format    *unused2    = workbook->add_format();
    xlsxwriter::format    *unused3    = workbook->add_format();


    /* Avoid warnings about unused variables since this test is checking
     * how unused formats are handled.
     */
    (void)worksheet2;
    (void)unused1;
    (void)unused2;
    (void)unused3;

    format->set_bold();

    worksheet1->write_string(0, 0, "Foo", NULL);
    worksheet1->write_number(1, 0, 123, NULL);

    worksheet3->write_string(1, 1, "Foo", NULL);
    worksheet3->write_string(2, 1, "Bar", format);
    worksheet3->write_number(3, 2, 234, NULL);


    /* For testing. This doesn't have a string or format and should be ignored. */
    worksheet1->write_string(0, 0, NULL, NULL);

    /* For testing. This doesn't have a formula and should be ignored. */
    worksheet1->write_formula(0, 0, NULL, NULL);

    int result = workbook->close(); return result;
}
