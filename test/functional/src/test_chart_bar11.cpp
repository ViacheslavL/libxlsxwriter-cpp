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

    std::shared_ptr<xlsxwriter::workbook> workbook = std::make_shared<xlsxwriter::workbook>("test_chart_bar11.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart1 = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);
    xlsxwriter::chart     *chart2 = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);
    xlsxwriter::chart     *chart3 = workbook->add_chart( xlsxwriter::LXW_CHART_BAR);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart1->axis_id_1 = 40274944;
    chart1->axis_id_2 = 40294272;
    chart2->axis_id_1 = 62355328;
    chart2->axis_id_2 = 62356864;
    chart3->axis_id_1 = 79538816;
    chart3->axis_id_2 = 65422464;


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

    worksheet->write_url(CELL("A7"), "http://www.perl.com/", NULL);
    worksheet->write_url(CELL("A8"), "http://www.perl.org/", NULL);
    worksheet->write_url(CELL("A9"), "http://www.perl.net/", NULL);

    chart1->add_series("", "=Sheet1!$A$1:$A$5");
    chart1->add_series("", "=Sheet1!$B$1:$B$5");
    chart1->add_series("", "=Sheet1!$C$1:$C$5");

    chart2->add_series("", "=Sheet1!$A$1:$A$5");
    chart2->add_series("", "=Sheet1!$B$1:$B$5");

    chart3->add_series("", "=Sheet1!$A$1:$A$5");

    worksheet->insert_chart(CELL("E9"), chart1);
    worksheet->insert_chart(CELL("D25"), chart2);
    worksheet->insert_chart(CELL("L32"), chart3);

    int result = workbook->close(); return result;
}
