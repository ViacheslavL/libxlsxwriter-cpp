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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image08.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    lxw_image_options options = {.x_scale = 0.5, .y_scale = 0.5};

    worksheet_insert_image_opt(worksheet, CELL("B3"), "images/grey.png", &options);

    int result = workbook->close(); return result;
}
