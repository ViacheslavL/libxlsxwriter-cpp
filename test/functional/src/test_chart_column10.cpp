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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_chart_column10.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_COLUMN);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 45686144;
    chart->axis_id_2 = 45722240;

    std::string data_1[5] = { "A", "B", "C", "D", "E"};
    uint8_t data_2[5] = {  1,   2,   3,   2,   1 };

    int row;
    for (row = 0; row < 5; row++) {
        worksheet->write_string(row, 0, data_1[row], NULL);
        worksheet->write_number(row, 1, data_2[row], NULL);
    }

    chart->add_series(
         "=Sheet1!$A$1:$A$5",
         "=Sheet1!$B$1:$B$5"
    );

    worksheet->insert_chart(CELL("E9"), chart);

    int result = workbook->close(); return result;
}
