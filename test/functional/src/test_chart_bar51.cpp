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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_chart_bar51.xlsx");
    xlsxwriter::worksheet    *worksheet = workbook->add_worksheet();
    xlsxwriter::chart        *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);
    xlsxwriter::chart_series *series1;
    xlsxwriter::chart_series *series2;
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

    series1 = chart->add_series("", "Sheet1!$A$1:$A$5");
    series2 = chart->add_series("", "Sheet1!$B$1:$B$5");


    /* Add the cached data for testing. */
    chart_add_data_cache(series1->values.get(), data[0], 5, 3, 0);
    chart_add_data_cache(series2->values.get(), data[0], 5, 3, 1);

    worksheet->insert_chart(CELL("E9"), chart);

    int result = workbook->close(); return result;
}
