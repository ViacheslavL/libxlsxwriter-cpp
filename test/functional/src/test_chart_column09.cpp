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

    std::shared_ptr<xlsxwriter::workbook> workbook = std::make_shared<xlsxwriter::workbook>("test_chart_column09.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_COLUMN);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 47400832;
    chart->axis_id_2 = 61387136;

    uint8_t data[5][2] = {
        {1, 1},
        {2, 2},
        {3, 3},
        {4, 2},
        {5, 1}
    };

    int row, col;
    for (row = 0; row < 5; row++)
        for (col = 0; col < 2; col++)
            worksheet->write_number(row, col, data[row][col] , NULL);

    chart->add_series(
                     "=Sheet1!$A$1:$A$5",
                     "=Sheet1!$B$1:$B$5");

    worksheet->insert_chart(CELL("E9"), chart);

    int result = workbook->close(); return result;
}
