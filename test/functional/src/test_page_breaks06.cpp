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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_page_breaks06.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    lxw_row_t hbreaks[] = {1, 5, 8, 13, 0};
    lxw_col_t vbreaks[] = {1, 3, 8, 0};

    worksheet->set_paper(9);
    worksheet->set_vertical_dpi(200);

    worksheet_set_h_pagebreaks(worksheet, hbreaks);
    worksheet_set_v_pagebreaks(worksheet, vbreaks);

    worksheet->write_string(CELL("A1"), "Foo" , NULL);

    int result = workbook->close(); return result;
}
