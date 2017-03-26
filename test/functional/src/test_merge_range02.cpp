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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_merge_range02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format *format = workbook->add_format();
    format_set_align(format, LXW_ALIGN_CENTER);

    worksheet_merge_range(worksheet, 1, 1, 5, 3, "Foo", format);

    int result = workbook->close(); return result;
}
