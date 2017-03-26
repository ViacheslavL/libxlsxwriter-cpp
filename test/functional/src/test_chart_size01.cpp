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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_chart_size01.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_COLUMN);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 61355904;
    chart->axis_id_2 = 61365248;

    uint8_t data[5][3] = {
        {1, 2,  3},
        {2, 4,  6},
        {3, 6,  9},
        {4, 8,  12},
        {5, 10, 15}
    };

    int row, col;
    for (row = 0; row < 5; row++)
        for (col = 0; col < 3; col++)
            worksheet->write_number(row, col, data[row][col], NULL);

    chart->add_series(NULL, "=Sheet1!$A$1:$A$5");
    chart->add_series(NULL, "=Sheet1!$B$1:$B$5");
    chart->add_series(NULL, "=Sheet1!$C$1:$C$5");

    xlsxwriter::image_options options = {};
    options.x_scale  = 1.06666667;
    options.y_scale  = 1.11111112;
    worksheet->insert_chart_opt(CELL("E9"), chart, &options);

    int result = workbook->close(); return result;
}
