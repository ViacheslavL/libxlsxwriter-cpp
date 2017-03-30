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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_defined_name02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet( "sheet One");

    workbook->define_name("Sales", "='sheet One'!$G$1:$H$10");

    (void)worksheet;

    int result = workbook->close(); return result;
}
