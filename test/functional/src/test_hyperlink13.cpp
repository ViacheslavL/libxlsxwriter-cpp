/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Test to compare output against Excel files.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink13.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format *format = workbook->add_format();

    format_set_align(format, LXW_ALIGN_CENTER);

    worksheet_merge_range(worksheet, RANGE("C4:E5"), "http://www.perl.org/", format);
    worksheet_write_url(worksheet, CELL("C4"), "http://www.perl.org/", format);

    int result = workbook->close(); return result;
}
