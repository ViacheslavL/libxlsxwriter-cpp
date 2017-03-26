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

    xlsxwriter::workbook       *workbook   = new xlsxwriter::workbook("test_properties02.xlsx");
    xlsxwriter::worksheet      *worksheet  = workbook->add_worksheet();
    xlsxwriter::doc_properties properties = {};

    properties.hyperlink_base = strdup("C:\\");

    workbook->set_properties(properties);

    (void)worksheet;
    int result = workbook->close(); return result;
}
