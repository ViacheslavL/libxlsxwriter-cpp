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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_format02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *format1    = workbook->add_format();
    xlsxwriter::format    *format2    = workbook->add_format();

    worksheet->set_row(0, 30, NULL);

    format_set_font_name(format1, "Arial");
    format_set_bold(format1);
    format_set_align(format1, LXW_ALIGN_LEFT);
    format_set_align(format1, LXW_ALIGN_VERTICAL_BOTTOM);

    format_set_font_name(format2, "Arial");
    format_set_bold(format2);
    format_set_rotation(format2, 90);
    format_set_align(format2, LXW_ALIGN_CENTER);
    format_set_align(format2, LXW_ALIGN_VERTICAL_BOTTOM);

    worksheet->write_string(0, 0, "Foo", format1);
    worksheet->write_string(0, 1, "Bar", format2);

    int result = workbook->close(); return result;
}
