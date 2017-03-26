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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_row_col_format08.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format    *bold      = workbook->add_format();
    bold->set_bold();

    xlsxwriter::format    *mixed     = workbook->add_format();
    format_set_bold(mixed);
    format_set_italic(mixed);

    xlsxwriter::format    *italic    = workbook->add_format();
    italic->set_italic();

    /* Manually force the format index order for testing. */
    workbook->set_default_xf_indices();

    worksheet->set_row(0, 15, bold);
    worksheet->set_column(0, 0, 8.43, italic);

    worksheet->write_string(0, 0, "Foo", mixed);
    worksheet->write_string(0, 1, "Foo", NULL);
    worksheet->write_string(1, 0, "Foo", NULL);
    worksheet->write_string(1, 1, "Foo", NULL);


    int result = workbook->close(); return result;
}
