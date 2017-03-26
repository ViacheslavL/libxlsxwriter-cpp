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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_gh42_01.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    char string[] = {0xe5, 0x9b, 0xbe, 0x14, 0xe5, 0x9b, 0xbe, 0x00};

    worksheet->write_string(0, 0, string, NULL);

    workbook_close(workbook);

    return 0;
}

