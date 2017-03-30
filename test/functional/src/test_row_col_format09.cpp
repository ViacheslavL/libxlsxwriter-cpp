/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case to test worksheet set_row() and set_column().
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_row_col_format09.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *bold      = workbook->add_format();
    bold->set_bold();

    xlsxwriter::format    *mixed     = workbook->add_format();
    mixed->set_bold();
    mixed->set_italic();

    xlsxwriter::format    *italic    = workbook->add_format();
    italic->set_italic();

    /* Manually force the format index order for testing. */
    workbook->set_default_xf_indices();

    worksheet->set_row(4, 15, bold);
    worksheet->set_column(2, 2, 8.43, italic);

    worksheet->write_string(0, 2, "Foo", NULL);
    worksheet->write_string(4, 0, "Foo", NULL);
    worksheet->write_string(4, 2, "Foo", mixed);

    int result = workbook->close(); return result;
}
