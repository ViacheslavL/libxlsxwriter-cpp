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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_image22.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->insert_image(CELL("B2"), "images/black_300.jpg");

    int result = workbook->close(); return result;
}
