/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case for TODO.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_landscape01.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_string(0, 0, "Foo" , NULL);
    worksheet->set_landscape();
    worksheet->set_paper(9);
    worksheet->set_vertical_dpi(200);

    int result = workbook->close(); return result;
}
