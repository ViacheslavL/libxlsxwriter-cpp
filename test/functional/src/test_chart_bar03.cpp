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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_chart_bar03.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart1    = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);
    xlsxwriter::chart     *chart2    = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart1->axis_id_1 = 64265216;
    chart1->axis_id_2 = 64447616;

    chart2->axis_id_1 = 86048128;
    chart2->axis_id_2 = 86058112;

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

    chart_add_series(chart1,
         "=Sheet1!$A$1:$A$5",
         "=Sheet1!$B$1:$B$5"
    );

    chart_add_series(chart1,
         "=Sheet1!$A$1:$A$5",
         "=Sheet1!$C$1:$C$5"
    );

    worksheet_insert_chart(worksheet, CELL("E9"), chart1);

    chart_add_series(chart2,
         "=Sheet1!$A$1:$A$4",
         "=Sheet1!$B$1:$B$4"
    );

    chart_add_series(chart2,
         "=Sheet1!$A$1:$A$4",
         "=Sheet1!$C$1:$C$4"
    );

    worksheet_insert_chart(worksheet, CELL("F25"), chart2);


    int result = workbook->close(); return result;
}
