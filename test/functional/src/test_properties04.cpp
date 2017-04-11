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

    std::shared_ptr<xlsxwriter::workbook> workbook = std::make_shared<xlsxwriter::workbook>("test_properties04.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    lxw_datetime   datetime  = {2016, 12, 12,  23, 0, 0};

    workbook->set_custom_property_string ( "Checked by",     "Adam");
    workbook->set_custom_property_datetime("Date completed",  &datetime);
    workbook->set_custom_property_integer ("Document number", 12345);
    workbook->set_custom_property_number  ("Reference",       1.2345);
    workbook->set_custom_property_boolean ("Source",          1);
    workbook->set_custom_property_boolean ("Status",          0);
    workbook->set_custom_property_string ( "Department",      "Finance");
    workbook->set_custom_property_number  ("Group",           1.2345678901234);

    worksheet->set_column(0, 0, 70, NULL);
    worksheet->write_string(CELL("A1"), "Select 'Office Button -> Prepare -> Properties' to see the file properties." , NULL);

    int result = workbook->close(); return result;
}
