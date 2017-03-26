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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_chart_axis02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_COLUMN);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 43704320;
    chart->axis_id_2 = 43706624;

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

    chart->get_x_axis()->set_name("XXX");
    chart->get_y_axis()->set_name("YYY");

    worksheet->insert_chart(CELL("E9"), chart);

    int result = workbook->close(); return result;
}
