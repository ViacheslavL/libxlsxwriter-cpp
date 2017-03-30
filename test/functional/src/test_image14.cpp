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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_image14.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->set_row(   1,     4.5,  NULL);
    worksheet->set_row(   2,    35.25, NULL);
    worksheet->set_column(2, 4,  3.29, NULL);
    worksheet->set_column(5, 5, 10.71, NULL);

    worksheet->insert_image(CELL("C2"), "images/logo.png");

    int result = workbook->close(); return result;
}
