/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Test to compare output against Excel files.
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook *workbook = new xlsxwriter::workbook("test_properties05.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    workbook->set_custom_property_string ( "Location", "Café");

    worksheet->set_column(0, 0, 70);
    worksheet->write_string(CELL("A1"), "Select 'Office Button -> Prepare -> Properties' to see the file properties.");

    int result = workbook->close(); return result;
}
