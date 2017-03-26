/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case for TODO.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_print_across01.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_print_across(worksheet);
    worksheet_set_paper(worksheet, 9);
    worksheet->vertical_dpi = 200;

    worksheet->write_string(0, 0, "Foo" , NULL);

    int result = workbook->close(); return result;
}
