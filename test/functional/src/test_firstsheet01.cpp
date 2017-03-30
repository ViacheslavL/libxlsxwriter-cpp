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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_firstsheet01.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet3 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet4 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet5 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet6 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet7 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet8 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet9 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet10 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet11 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet12 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet13 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet14 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet15 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet16 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet17 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet18 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet19 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet20 = workbook->add_worksheet();

    worksheet8->set_first_sheet();
    worksheet20->activate();

    /* Avoid warnings about unused variables. */
    (void)worksheet1;
    (void)worksheet2;
    (void)worksheet3;
    (void)worksheet4;
    (void)worksheet5;
    (void)worksheet6;
    (void)worksheet7;
    (void)worksheet9;
    (void)worksheet10;
    (void)worksheet11;
    (void)worksheet12;
    (void)worksheet13;
    (void)worksheet14;
    (void)worksheet15;
    (void)worksheet16;
    (void)worksheet17;
    (void)worksheet18;
    (void)worksheet19;

    int result = workbook->close(); return result;
}
