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

    std::shared_ptr<xlsxwriter::workbook> workbook = std::make_shared<xlsxwriter::workbook>("test_row_col_format14.xlsx");
    xlsxwriter::worksheet *worksheet    = workbook->add_worksheet();
    xlsxwriter::row_col_options options = {1, 0, 0};
    xlsxwriter::format    *bold         = workbook->add_format();

    bold->set_bold();

    worksheet->set_column(1, 3, 5, NULL);
    worksheet->set_column(5, 5, 8, NULL);
    worksheet->set_column(7, 7, LXW_DEF_COL_WIDTH, bold);
    worksheet->set_column(9, 9, 2, NULL);
    worksheet->set_column_opt(11, 11, LXW_DEF_COL_WIDTH, NULL, options);

    int result = workbook->close(); return result;
}
