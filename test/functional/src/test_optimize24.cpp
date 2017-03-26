/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Test case for writing data in optimization mode.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook_options options = {1, NULL};

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_optimize24.xlsx", &options);
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format    *bold      = workbook->add_format();

    format_set_bold(bold);

    worksheet->set_row(0, 20, bold);
    worksheet->write_string(0, 0, "Foo", NULL);

    int result = workbook->close(); return result;
}
