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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_simple03.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet("Data Sheet");
    xlsxwriter::worksheet *worksheet3 = workbook->add_worksheet();

    xlsxwriter::format *bold = workbook->add_format();
    bold->set_bold();

    worksheet1->write_string(CELL("A1"), "Foo" , NULL);
    worksheet1->write_number(CELL("A2"), 123 , NULL);

    worksheet3->write_string(CELL("B2"), "Foo" , NULL);
    worksheet3->write_string(CELL("B3"), "Bar", bold);
    worksheet3->write_number(CELL("C4"), 234 , NULL);

    /* Ensure the active worksheet is overwritten, below. */
    worksheet2->activate();

    worksheet2->select();
    worksheet3->select();
    worksheet3->activate();

    int result = workbook->close(); return result;
}
