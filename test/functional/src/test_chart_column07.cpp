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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_chart_column07.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_COLUMN);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 68810240;
    chart->axis_id_2 = 68811776;

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

    xlsxwriter::chart_series *series1 = chart->add_series("", "=(Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5)");

    worksheet->insert_chart(CELL("E9"), chart);


    /* Add the cached data for testing. */
    uint8_t test_data[4][3] = {
        {1, 2,  3},
        {2, 4,  6},
        {4, 8,  12},
        {5, 10, 15}
    };

    chart_add_data_cache(series1->values.get(), test_data[0], 4, 3, 0);


    int result = workbook->close(); return result;
}
