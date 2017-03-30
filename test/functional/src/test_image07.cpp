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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_image07.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();

    worksheet1->insert_image(CELL("E9"), "images/red.png");
    worksheet2->insert_image(CELL("E9"), "images/yellow.png");

    int result = workbook->close(); return result;
}
