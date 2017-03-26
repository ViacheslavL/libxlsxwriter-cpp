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

    xlsxwriter::image_options options = {};
    options.x_offset = 8;
    options.y_offset = 5;

    worksheet->set_row(1, 75, NULL);
    worksheet->set_column(2, 2, 32, NULL);

    worksheet->insert_image_opt(CELL("C2"), "images/logo.png", &options);

    int result = workbook->close(); return result;
}
