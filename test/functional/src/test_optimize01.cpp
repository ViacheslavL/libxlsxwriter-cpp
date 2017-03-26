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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_optimize01.xlsx", &options);
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_string(0, 0, "Hello", NULL);
    /* For testing overwrite the 0, 0 cell to ensure the original is freed. */
    worksheet->write_string(0, 0, "Hello", NULL);

    worksheet->write_number(1, 0, 123,     NULL);

    int result = workbook->close(); return result;
}
