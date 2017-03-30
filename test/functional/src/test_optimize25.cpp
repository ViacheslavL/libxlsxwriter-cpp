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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_optimize25.xlsx", options);
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format    *bold      = workbook->add_format();

    bold->set_bold();

    worksheet->set_row(0, 20, bold);
    worksheet->write_string(2, 0, "Foo", NULL);

    int result = workbook->close(); return result;
}
