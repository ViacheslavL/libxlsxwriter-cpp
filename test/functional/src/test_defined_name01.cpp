/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case for defined names.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_defined_name01.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet3 = workbook->add_worksheet("Sheet 3");

    worksheet1->set_paper(9);
    worksheet1->set_vertical_dpi(200);

    worksheet1->print_area(RANGE("A1:E6"));
    worksheet1->autofilter(RANGE("F1:G1"));
    worksheet1->write_string(CELL("G1"), "Filter", NULL);
    worksheet1->write_string(CELL("F1"), "Auto", NULL);
    worksheet1->fit_to_pages(2, 2);

    workbook->define_name("'Sheet 3'!Bar", "='Sheet 3'!$A$1");
    workbook->define_name("Abc",           "=Sheet1!$A$1");
    workbook->define_name("Baz",           "=0.98");
    workbook->define_name("Sheet1!Bar",    "=Sheet1!$A$1");
    workbook->define_name("Sheet2!Bar",    "=Sheet2!$A$1");
    workbook->define_name("Sheet2!aaa",    "=Sheet2!$A$1");
    workbook->define_name("_Egg",          "=Sheet1!$A$1");
    workbook->define_name("_Fog",          "=Sheet1!$A$1");

    (void)worksheet2;
    (void)worksheet3;

    int result = workbook->close(); return result;
}
