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

    xlsxwriter::workbook *workbook   = new_workbook("test_chart_order01.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet3 = workbook->add_worksheet();
    xlsxwriter::chart     *chart1     = workbook->add_chart( xlsxwriter::LXW_CHART_COLUMN);
    xlsxwriter::chart     *chart2     = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);
    xlsxwriter::chart     *chart3     = workbook->add_chart( xlsxwriter::LXW_CHART_LINE);
    xlsxwriter::chart     *chart4     = workbook->add_chart( xlsxwriter::LXW_CHART_PIE);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart1->axis_id_1 = 54976896;
    chart1->axis_id_2 = 54978432;

    chart2->axis_id_1 = 54310784;
    chart2->axis_id_2 = 54312320;

    chart3->axis_id_1 = 69816704;
    chart3->axis_id_2 = 69818240;

    chart4->axis_id_1 = 69816704;
    chart4->axis_id_2 = 69818240;

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
            worksheet_write_number(worksheet1, row, col, data[row][col], NULL);
            worksheet_write_number(worksheet2, row, col, data[row][col], NULL);
            worksheet_write_number(worksheet3, row, col, data[row][col], NULL);
        }

    chart_add_series(chart1, NULL, "=Sheet1!$A$1:$A$5");
    chart_add_series(chart2, NULL, "=Sheet2!$A$1:$A$5");
    chart_add_series(chart3, NULL, "=Sheet3!$A$1:$A$5");
    chart_add_series(chart4, NULL, "=Sheet1!$B$1:$B$5");

    worksheet_insert_chart(worksheet1, CELL("E9"),  chart1);
    worksheet_insert_chart(worksheet2, CELL("E9"),  chart2);
    worksheet_insert_chart(worksheet3, CELL("E9"),  chart3);
    worksheet_insert_chart(worksheet1, CELL("E24"), chart4);

    int result = workbook->close(); return result;
}
