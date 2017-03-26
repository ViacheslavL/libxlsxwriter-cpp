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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_protect03.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::format *unlocked = workbook->add_format();
    format_set_unlocked(unlocked);

    xlsxwriter::format *hidden = workbook->add_format();
    format_set_unlocked(hidden);
    format_set_hidden(hidden);

    worksheet_protect(worksheet, "password", NULL);

    worksheet->write_number(CELL("A1"), 1 , NULL);
    worksheet->write_number(CELL("A2"), 2, unlocked);
    worksheet->write_number(CELL("A3"), 3, hidden);

    int result = workbook->close(); return result;
}
