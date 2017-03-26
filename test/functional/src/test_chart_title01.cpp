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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_chart_title01.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_COLUMN);

    xlsxwriter::chart_series *series;

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 46165376;
    chart->axis_id_2 = 54462720;

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

    series = chart->add_series(NULL, "=Sheet1!$A$1:$A$5");

    series->set_name("Foo");
    chart->title_off();

    worksheet->insert_chart(CELL("E9"), chart);

    int result = workbook->close(); return result;
}
