/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case fort set_page_view().
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_page_view01.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet_set_page_view(worksheet);

    worksheet->write_string(CELL("A1"), "Foo" , NULL);

    worksheet_set_paper(worksheet, 9);
    worksheet->vertical_dpi = 200;

    int result = workbook->close(); return result;
}
