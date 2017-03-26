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

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_chart_bar22.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    lxw_chart     *chart     = workbook_add_chart(workbook, LXW_CHART_BAR);

    /* For testing, copy the randomly generated axis ids in the target file. */
    chart->axis_id_1 = 43706240;
    chart->axis_id_2 = 43727104;


    worksheet->write_string(0, 1, "Series 1", NULL);
    worksheet->write_string(0, 2, "Series 2", NULL);
    worksheet->write_string(0, 3, "Series 3", NULL);

    worksheet->write_string(1, 0, "Category 1", NULL);
    worksheet->write_string(2, 0, "Category 2", NULL);
    worksheet->write_string(3, 0, "Category 3", NULL);
    worksheet->write_string(4, 0, "Category 4", NULL);

    worksheet->write_number(1, 1, 4.3, NULL);
    worksheet->write_number(2, 1, 2.5, NULL);
    worksheet->write_number(3, 1, 3.5, NULL);
    worksheet->write_number(4, 1, 4.5, NULL);

    worksheet->write_number(1, 2, 2.4, NULL);
    worksheet->write_number(2, 2, 4.5, NULL);
    worksheet->write_number(3, 2, 1.8, NULL);
    worksheet->write_number(4, 2, 2.8, NULL);

    worksheet->write_number(1, 3, 2,   NULL);
    worksheet->write_number(2, 3, 2,   NULL);
    worksheet->write_number(3, 3, 3,   NULL);
    worksheet->write_number(4, 3, 5,   NULL);

    worksheet->set_column(COLS("A:D"), 11, NULL);

    chart_add_series(chart,
         "=Sheet1!$A$2:$A$5",
         "=Sheet1!$B$2:$B$5"
    );

    chart_add_series(chart,
         "=Sheet1!$A$2:$A$5",
         "=Sheet1!$C$2:$C$5"
    );

    chart_add_series(chart,
         "=Sheet1!$A$2:$A$5",
         "=Sheet1!$D$2:$D$5"
    );

    worksheet_insert_chart(worksheet, CELL("E9"), chart);

    int result = workbook->close(); return result;
}
