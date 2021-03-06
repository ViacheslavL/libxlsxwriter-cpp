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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hyperlink13.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format *format = workbook->add_format();

    format->set_align(xlsxwriter::LXW_ALIGN_CENTER);

    worksheet->merge_range(RANGE("C4:E5"), "http://www.perl.org/", format);
    worksheet->write_url(CELL("C4"), "http://www.perl.org/", format);

    int result = workbook->close(); return result;
}
