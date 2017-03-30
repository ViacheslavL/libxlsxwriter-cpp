/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Simple test case to test data writing.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_data04.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_string(0,       0, "Foo",  NULL);
    worksheet->write_string(0,       1, "Bar",  NULL);
    worksheet->write_string(1,       0, "Bing", NULL);
    worksheet->write_string(2,       0, "Buzz", NULL);
    worksheet->write_string(1048575, 0, "End",  NULL);

    int result = workbook->close(); return result;
}
