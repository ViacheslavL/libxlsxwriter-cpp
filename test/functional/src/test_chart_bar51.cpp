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

    lxw_workbook     *workbook  = new_workbook("test_chart_bar51.xlsx");
    lxw_worksheet    *worksheet = workbook->add_worksheet();
    xlsxwriter::chart        *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);
    xlsxwriter::LXW_CHART_series *series1;
    xlsxwriter::LXW_CHART_series *series2;
    int row, col;

    uint8_t data[5][3] = {
        {1, 2,  3},
        {2, 4,  6},
        {3, 6,  9},
        {4, 8,  12},
        {5, 10, 15}
    };

    for (row = 0; row < 5; row++)
        for (col = 0; col < 3; col++)
            worksheet->write_number(row, col, data[row][col] , NULL);

    series1 = chart->add_series(NULL, "Sheet1!$A$1:$A$5");
    series2 = chart->add_series(NULL, "Sheet1!$B$1:$B$5");


    /* Add the cached data for testing. */
    xlsxwriter::LXW_CHART_add_data_cache(series1->values, data[0], 5, 3, 0);
    xlsxwriter::LXW_CHART_add_data_cache(series2->values, data[0], 5, 3, 1);

    worksheet_insert_chart(worksheet, CELL("E9"), chart);

    int result = workbook->close(); return result;
}
