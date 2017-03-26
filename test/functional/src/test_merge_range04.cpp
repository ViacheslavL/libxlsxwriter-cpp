/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case for merged ranges.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_merge_range04.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format *format = workbook->add_format();
    format->set_align(xlsxwriter::LXW_ALIGN_CENTER);
    format_set_bold(format);

    worksheet->merge_range(1, 1, 1, 3, "Foo", format);

    int result = workbook->close(); return result;
}
