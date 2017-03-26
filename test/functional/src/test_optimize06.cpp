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

    xlsxwriter::workbook_options options = {};
    options.constant_memory = true;

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_optimize06.xlsx", options);

    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    uint8_t i;
    char c[] = {0x00, 0x00};

    worksheet->write_string(0, 0, "_x0000_", NULL);

    for (i = 1; i <= 127; i++) {
        (*c)++;
        if (i != 34) {
            worksheet->write_string(i, 0, c, NULL);

        }
    }

    int result = workbook->close(); return result;
}
