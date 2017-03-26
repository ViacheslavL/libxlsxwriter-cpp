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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_chart_bar65.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 64264064;
    chart->axis_id_2 = 64447232;

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
            worksheet->write_number(row, col, data[row][col] , NULL);

    xlsxwriter::LXW_CHART_series *series1 = chart->add_series(NULL, "=Sheet1!$A$1:$A$5");
    xlsxwriter::LXW_CHART_series *series2 = chart->add_series(NULL, NULL);
    xlsxwriter::LXW_CHART_series *series3 = chart->add_series(NULL, NULL);

    chart_series_set_values(series1, "Sheet1", 0, 0, 4, 0);
    chart_series_set_values(series2, "Sheet1", 0, 1, 4, 1);
    chart_series_set_values(series3, "Sheet1", 0, 2, 4, 2);

    worksheet_insert_chart(worksheet, CELL("E9"), chart);

    int result = workbook->close(); return result;
}
