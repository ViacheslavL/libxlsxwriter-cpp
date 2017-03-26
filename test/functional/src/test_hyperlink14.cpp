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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink14.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format *format = workbook->add_format();

    format->set_align(xlsxwriter::LXW_ALIGN_CENTER);

    worksheet->write_string(CELL("A1"), "Perl Home", NULL);

    worksheet->merge_range(RANGE("C4:E5"), "http://www.perl.org/", format);
    worksheet_write_url_opt(worksheet, CELL("C4"), "http://www.perl.org/", format, "Perl Home", NULL);


    int result = workbook->close(); return result;
}
