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

    std::shared_ptr<xlsxwriter::workbook> workbook = std::make_shared<xlsxwriter::workbook>("test_chart_pie05.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_PIE);

    uint8_t data[3][2] = {
        {2,  60},
        {4,  30},
        {6,  10},
    };

    int row, col;
    for (row = 0; row < 3; row++)
        for (col = 0; col < 2; col++)
            worksheet->write_number(row, col, data[row][col], NULL);

    chart->add_series(
         "=Sheet1!$A$1:$A$3",
         "=Sheet1!$B$1:$B$3"
    );

    chart->set_rotation(45);

    worksheet->insert_chart(CELL("E9"), chart);

    int result = workbook->close(); return result;
}
