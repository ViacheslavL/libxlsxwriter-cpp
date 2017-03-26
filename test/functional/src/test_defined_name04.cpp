/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case for defined names.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_defined_name04.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    workbook_define_name(workbook, "\\__",     "=Sheet1!$A$1");
    workbook_define_name(workbook, "a3f6",     "=Sheet1!$A$2");
    workbook_define_name(workbook, "afoo.bar", "=Sheet1!$A$3");
    workbook_define_name(workbook, "étude",    "=Sheet1!$A$4");
    workbook_define_name(workbook, "eésumé",   "=Sheet1!$A$5");
    workbook_define_name(workbook, "a",        "=Sheet1!$A$6");

    (void)worksheet;

    int result = workbook->close(); return result;
}
