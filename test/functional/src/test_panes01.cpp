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

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_panes01.xlsx");
    xlsxwriter::worksheet *worksheet01 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet02 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet03 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet04 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet05 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet06 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet07 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet08 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet09 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet10 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet11 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet12 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet13 = workbook->add_worksheet();

    worksheet_write_string(worksheet01, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet02, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet03, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet04, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet05, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet06, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet07, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet08, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet09, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet10, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet11, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet12, CELL("A1"), "Foo" , NULL);
    worksheet_write_string(worksheet13, CELL("A1"), "Foo" , NULL);

    worksheet_freeze_panes(worksheet01, CELL("A2"));
    worksheet_freeze_panes(worksheet02, CELL("A3"));
    worksheet_freeze_panes(worksheet03, CELL("B1"));
    worksheet_freeze_panes(worksheet04, CELL("C1"));
    worksheet_freeze_panes(worksheet05, CELL("B2"));
    worksheet_freeze_panes(worksheet06, CELL("G4"));
    worksheet_freeze_panes_opt(worksheet07, 3, 6, 3, 6, 1);
    worksheet_split_panes(worksheet08, 15, 0);
    worksheet_split_panes(worksheet09, 30, 0);
    worksheet_split_panes(worksheet10, 0, 8.46);
    worksheet_split_panes(worksheet11, 0, 17.57);
    worksheet_split_panes(worksheet12, 15, 8.46);
    worksheet_split_panes(worksheet13, 45, 54.14);

    int result = workbook->close(); return result;
}
