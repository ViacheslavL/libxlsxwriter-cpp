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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image29.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    lxw_image_options options = {.x_offset = -210, .y_offset = 1};

    worksheet_insert_image_opt(worksheet, 0, 10, "images/red_208.png", &options);

    int result = workbook->close(); return result;
}
