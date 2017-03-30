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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_page_breaks06.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    std::vector<lxw_row_t> hbreaks = {1, 5, 8, 13};
    std::vector<lxw_col_t> vbreaks = {1, 3, 8};

    worksheet->set_paper(9);
    worksheet->set_vertical_dpi(200);

    worksheet->set_h_pagebreaks(hbreaks);
    worksheet->set_v_pagebreaks(vbreaks);

    worksheet->write_string(CELL("A1"), "Foo" , NULL);

    int result = workbook->close(); return result;
}
