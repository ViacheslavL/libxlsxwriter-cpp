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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_image15.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    xlsxwriter::image_options options = {.x_offset = 13, .y_offset = 2};

    worksheet->set_row(   1,     4.5,  NULL);
    worksheet->set_row(   2,    35.25, NULL);
    worksheet->set_column(2, 4,  3.29, NULL);
    worksheet->set_column(5, 5, 10.71, NULL);

    worksheet->insert_image_opt(CELL("C2"), "images/logo.png", &options);

    int result = workbook->close(); return result;
}
