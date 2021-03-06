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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_image33.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::image_options options = {};
    options.x_offset = -2;
    options.y_offset = -1;

    worksheet->set_column(3, 3, 3.86, NULL);
    worksheet->set_column(4, 4, 1.43, NULL);
    worksheet->set_row(7, 7.5, NULL);
    worksheet->set_row(8, 9.75, NULL);

    worksheet->insert_image_opt(CELL("E9"), "images/red.png", &options);

    int result = workbook->close(); return result;
}
