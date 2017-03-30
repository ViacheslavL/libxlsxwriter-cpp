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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_image08.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::image_options options = {};
    options.x_scale = 0.5;
    options.y_scale = 0.5;

    worksheet->insert_image_opt(CELL("B3"), "images/grey.png", &options);

    int result = workbook->close(); return result;
}
