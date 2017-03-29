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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_page_breaks02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    std::vector<lxw_row_t> breaks = {3, 7, 15};

    worksheet->set_paper(9);
    worksheet->set_vertical_dpi(200);

    worksheet->set_h_pagebreaks(breaks);

    worksheet->write_string(CELL("A1"), "Foo" , NULL);

    int result = workbook->close(); return result;
}
