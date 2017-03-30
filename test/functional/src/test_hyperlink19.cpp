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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hyperlink19.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    /* This test requires that we check if the cell that the hyperlink refers
     * to is a string. If it isn't be have to add a display attribute. However,
     * looking up the cell is currently too expensive.
     */ 
    worksheet->write_url(CELL("A1"), "http://www.perl.com/", NULL);
    worksheet->write_formula_num(CELL("A1"), "=1+1", NULL, 2);

    workbook->sst->string_count = 0;

    int result = workbook->close(); return result;
}
