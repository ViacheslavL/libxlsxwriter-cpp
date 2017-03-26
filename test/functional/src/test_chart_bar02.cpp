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

    xlsxwriter::workbook *workbook   = new xlsxwriter::workbook("test_chart_bar02.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();
    xlsxwriter::chart     *chart      = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 93218304;
    chart->axis_id_2 = 93219840;

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
            worksheet2->write_number(row, col, data[row][col] , NULL);

    worksheet1->write_string(CELL("A1"), "Foo" , NULL);

    chart->add_series("Sheet2!$A$1:$A$5", "Sheet2!$B$1:$B$5");
    chart->add_series("Sheet2!$A$1:$A$5", "Sheet2!$C$1:$C$5");


    worksheet2->insert_chart(CELL("E9"), chart);

    int result = workbook->close(); return result;
}
