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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_merge_range03.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format *format = workbook->add_format();
    format->set_align(xlsxwriter::LXW_ALIGN_CENTER);

    worksheet->merge_range(1, 1, 1, 2, "Foo", format);
    worksheet->merge_range(1, 3, 1, 4, "Foo", format);
    worksheet->merge_range(1, 5, 1, 6, "Foo", format);

    int result = workbook->close(); return result;
}
