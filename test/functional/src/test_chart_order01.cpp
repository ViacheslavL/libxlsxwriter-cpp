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

    xlsxwriter::workbook *workbook   = new xlsxwriter::workbook("test_chart_order01.xlsx");
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
            worksheet1->write_number(row, col, data[row][col], NULL);
            worksheet2->write_number(row, col, data[row][col], NULL);
            worksheet3->write_number(row, col, data[row][col], NULL);
        }

    chart1->add_series("", "=Sheet1!$A$1:$A$5");
    chart2->add_series("", "=Sheet2!$A$1:$A$5");
    chart3->add_series("", "=Sheet3!$A$1:$A$5");
    chart4->add_series("", "=Sheet1!$B$1:$B$5");

    worksheet1->insert_chart(CELL("E9"),  chart1);
    worksheet2->insert_chart(CELL("E9"),  chart2);
    worksheet3->insert_chart(CELL("E9"),  chart3);
    worksheet1->insert_chart(CELL("E24"), chart4);

    int result = workbook->close(); return result;
}
