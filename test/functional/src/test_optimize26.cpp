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

    /* Use deprecated constructor for testing. */
    std::shared_ptr<xlsxwriter::workbook> workbook = std::make_shared<xlsxwriter::workbook>("test_optimize26.xlsx", options);
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_string(2, 2, "café", NULL);

    int result = workbook->close(); return result;
}
