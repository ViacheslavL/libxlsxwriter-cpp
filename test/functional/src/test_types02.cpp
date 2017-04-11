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

    std::shared_ptr<xlsxwriter::workbook> workbook = std::make_shared<xlsxwriter::workbook>("test_types02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_boolean(CELL("A1"), 1 , NULL);
    worksheet->write_boolean(CELL("A2"), 0 , NULL);

    int result = workbook->close(); return result;
}
