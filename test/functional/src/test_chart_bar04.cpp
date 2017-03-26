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

    xlsxwriter::workbook *workbook   = new xlsxwriter::workbook("test_chart_bar04.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();
    xlsxwriter::chart     *chart1     = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);
    xlsxwriter::chart     *chart2     = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart1->axis_id_1 = 64446848;
    chart1->axis_id_2 = 64448384;

    chart2->axis_id_1 = 85389696;
    chart2->axis_id_2 = 85391232;

    uint8_t data[5][3] = {
        {1, 2,  3},
        {2, 4,  6},
        {3, 6,  9},
        {4, 8,  12},
        {5, 10, 15}
    };

    int row, col;
    for (row = 0; row < 5; row++)
        for (col = 0; col < 3; col++) {
            worksheet1->write_number(row, col, data[row][col] , NULL);
            worksheet2->write_number(row, col, data[row][col] , NULL);
        }


    chart1->add_series(
         "=Sheet1!$A$1:$A$5",
         "=Sheet1!$B$1:$B$5"
    );

    chart1->add_series(
         "=Sheet1!$A$1:$A$5",
         "=Sheet1!$C$1:$C$5"
    );

    worksheet1->insert_chart(CELL("E9"), chart1);


    chart2->add_series(
         "=Sheet2!$A$1:$A$5",
         "=Sheet2!$B$1:$B$5"
    );

    chart2->add_series(
         "=Sheet2!$A$1:$A$5",
         "=Sheet2!$C$1:$C$5"
    );

    worksheet2->insert_chart(CELL("E9"), chart2);

    int result = workbook->close(); return result;
}
