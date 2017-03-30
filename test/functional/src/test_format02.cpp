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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_format02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *format1    = workbook->add_format();
    xlsxwriter::format    *format2    = workbook->add_format();

    worksheet->set_row(0, 30, NULL);

    format1->set_font_name("Arial");
    format1->set_bold();
    format1->set_align(xlsxwriter::LXW_ALIGN_LEFT);
    format1->set_align(xlsxwriter::LXW_ALIGN_VERTICAL_BOTTOM);

    format2->set_font_name("Arial");
    format2->set_bold();
    format2->set_rotation(90);
    format2->set_align(xlsxwriter::LXW_ALIGN_CENTER);
    format2->set_align(xlsxwriter::LXW_ALIGN_VERTICAL_BOTTOM);

    worksheet->write_string(0, 0, "Foo", format1);
    worksheet->write_string(0, 1, "Bar", format2);

    int result = workbook->close(); return result;
}
