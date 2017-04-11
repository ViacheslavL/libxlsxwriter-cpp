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

    std::shared_ptr<xlsxwriter::workbook> workbook = std::make_shared<xlsxwriter::workbook>("test_chart_column05.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet( "Foo");
    xlsxwriter::chart     *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_COLUMN);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 47292800;
    chart->axis_id_2 = 47295104;

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

    chart->add_series("", "=Foo!$A$1:$A$5");
    chart->add_series("", "=Foo!$B$1:$B$5");
    chart->add_series("", "=Foo!$C$1:$C$5");

    worksheet->insert_chart(CELL("E9"), chart);

    int result = workbook->close(); return result;
}
