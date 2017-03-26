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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_chart_str01.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::chart     *chart     = workbook->add_chart( xlsxwriter::LXW_CHART_LINE);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 40501632;
    chart->axis_id_2 = 40514688;

    worksheet->write_number(0, 0, 1,     NULL);
    worksheet->write_number(1, 0, 2,     NULL);
    worksheet->write_number(2, 0, 3,     NULL);
    worksheet->write_number(3, 0, 4,     NULL);
    worksheet->write_number(4, 0, 5,     NULL);

    worksheet->write_number(0, 1, 2,     NULL);
    worksheet->write_number(1, 1, 4,     NULL);
    worksheet->write_string(2, 1, "Foo", NULL);
    worksheet->write_number(3, 1, 8,     NULL);
    worksheet->write_number(4, 1, 10,    NULL);

    worksheet->write_number(0, 2, 3,     NULL);
    worksheet->write_number(1, 2, 6,     NULL);
    worksheet->write_number(2, 2, 9,     NULL);
    worksheet->write_number(3, 2, 12,    NULL);
    worksheet->write_number(4, 2, 15,    NULL);

    worksheet->write_string(CELL("A6"), "Foo" , NULL);

    chart->add_series(NULL, "=Sheet1!$A$1:$A$5");
    chart->add_series(NULL, "=Sheet1!$B$1:$B$5");
    chart->add_series(NULL, "=Sheet1!$C$1:$C$5");

    worksheet_insert_chart(worksheet, CELL("E9"), chart);

    int result = workbook->close(); return result;
}
