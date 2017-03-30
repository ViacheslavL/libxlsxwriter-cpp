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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_defined_name04.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    workbook->define_name("\\__",     "=Sheet1!$A$1");
    workbook->define_name("a3f6",     "=Sheet1!$A$2");
    workbook->define_name("afoo.bar", "=Sheet1!$A$3");
    workbook->define_name("étude",    "=Sheet1!$A$4");
    workbook->define_name("eésumé",   "=Sheet1!$A$5");
    workbook->define_name("a",        "=Sheet1!$A$6");

    (void)worksheet;

    int result = workbook->close(); return result;
}
