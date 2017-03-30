/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case to test data writing.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook_options options;
    options.constant_memory = false;
    options.tmpdir = ".";

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_tmpdir01.xlsx", options);
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_string(0, 0, "Hello", NULL);
    worksheet->write_number(1, 0, 123,     NULL);

    int result = workbook->close(); return result;
}
