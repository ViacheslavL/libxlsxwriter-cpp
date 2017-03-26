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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_chart_doughnut05.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    lxw_chart     *chart     = workbook_add_chart(workbook, LXW_CHART_DOUGHNUT);

    uint8_t data[3][2] = {
        {2,  60},
        {4,  30},
        {6,  10},
    };

    int row, col;
    for (row = 0; row < 3; row++)
        for (col = 0; col < 2; col++)
            worksheet->write_number(row, col, data[row][col], NULL);

    chart_add_series(chart,
         "=Sheet1!$A$1:$A$3",
         "=Sheet1!$B$1:$B$3"
    );

    chart_set_rotation(chart, 360);

    worksheet_insert_chart(worksheet, CELL("E9"), chart);

    int result = workbook->close(); return result;
}
