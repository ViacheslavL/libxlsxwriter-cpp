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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_types08.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format *bold = workbook->add_format();
    bold->set_bold();

    xlsxwriter::format *italic = workbook->add_format();
    italic->set_italic();

    worksheet->write_boolean(CELL("A1"), 2, bold);
    worksheet->write_boolean(CELL("A2"), 0, italic);

    return workbook->close();
}
