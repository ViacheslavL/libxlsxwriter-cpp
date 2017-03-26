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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image13.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    lxw_image_options options = {.x_offset = 8, .y_offset = 5};

    worksheet->set_row(1, 75, NULL);
    worksheet->set_column(2, 2, 32, NULL);

    worksheet_insert_image_opt(worksheet, CELL("C2"), "images/logo.png", &options);

    int result = workbook->close(); return result;
}
