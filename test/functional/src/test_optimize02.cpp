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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_optimize02.xlsx", options);
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_string(0, 0, "Hello", NULL);
    worksheet->write_number(1, 0, 123,     NULL);

    /* G1 should be ignored since a later row has already been written. */
    worksheet->write_string(CELL("G1"), "World", NULL);

    int result = workbook->close(); return result;
}
