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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_hyperlink04.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet3 = workbook->add_worksheet("Data Sheet");

    (void)worksheet2;
    (void)worksheet3;

    worksheet1->write_url_opt(CELL("A1"),  "internal:Sheet2!A1",       NULL, NULL,        NULL);
    worksheet1->write_url_opt(CELL("A3"),  "internal:Sheet2!A1:A5",    NULL, NULL,        NULL);
    worksheet1->write_url_opt(CELL("A5"),  "internal:'Data Sheet'!D5", NULL, "Some text", NULL);
    worksheet1->write_url_opt(CELL("E12"), "internal:Sheet1!J1",       NULL, NULL,        NULL);
    worksheet1->write_url_opt(CELL("G17"), "internal:Sheet2!A1",       NULL, "Some text", NULL);
    worksheet1->write_url_opt(CELL("A18"), "internal:Sheet2!A1",       NULL, NULL,        "Tool Tip 1");
    worksheet1->write_url_opt(CELL("A20"), "internal:Sheet2!A1",       NULL, "More text", "Tool Tip 2");

    int result = workbook->close(); return result;
}
