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

    xlsxwriter::workbook_options options = {};
    options.constant_memory = true;

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_optimize22.xlsx", options);
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format    *bold      = workbook->add_format();

    bold->set_bold();

    worksheet->set_column(0, 0, 36, bold);

    int result = workbook->close(); return result;
}
