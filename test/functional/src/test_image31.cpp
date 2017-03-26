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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image31.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    lxw_image_options options = {.x_offset = -2, .y_offset = -1};

    worksheet->set_column(3, 3, 3.86, NULL);
    worksheet->set_row(7, 7.5, NULL);

    worksheet_insert_image_opt(worksheet, CELL("E9"), "images/red.png", &options);

    int result = workbook->close(); return result;
}
