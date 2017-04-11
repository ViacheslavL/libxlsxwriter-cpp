/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case to test worksheet set_row() and set_column().
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    std::shared_ptr<xlsxwriter::workbook>  workbook     = std::make_shared<xlsxwriter::workbook>("test_row_col_format13.xlsx");
    xlsxwriter::worksheet *worksheet    = workbook->add_worksheet();
    xlsxwriter::row_col_options options = {1, 0, 0};
    xlsxwriter::format    *bold         = workbook->add_format();

    bold->set_bold();

    worksheet->set_column(COLS("B:D"), 5, NULL);
    worksheet->set_column_opt(COLS("F:F"), 8, NULL, options);
    worksheet->set_column(COLS("H:H"), LXW_DEF_COL_WIDTH, bold);
    worksheet->set_column(COLS("J:J"), 2, NULL);
    worksheet->set_column_opt(COLS("L:L"), LXW_DEF_COL_WIDTH, NULL, options);

    int result = workbook->close(); return result;
}
