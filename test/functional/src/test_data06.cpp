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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_data06.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format    *format1   = workbook->add_format();
    xlsxwriter::format    *format2   = workbook->add_format();
    xlsxwriter::format    *format3   = workbook->add_format();

    format1->set_bold();

    format2->set_italic();

    format3->set_bold();
    format3->set_italic();

    worksheet->write_string(CELL("A1"), "Foo", format1);
    worksheet->write_string(CELL("A2"), "Bar", format2);
    worksheet->write_string(CELL("A3"), "Baz", format3);

    int result = workbook->close(); return result;
}

