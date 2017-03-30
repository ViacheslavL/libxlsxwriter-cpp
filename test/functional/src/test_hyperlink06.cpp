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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hyperlink06.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_url_opt(CELL("A1"), "external:C:\\Temp\\foo.xlsx",            NULL);
    worksheet->write_url_opt(CELL("A3"), "external:C:\\Temp\\foo.xlsx#Sheet1!A1",  NULL);
    worksheet->write_url_opt(CELL("A5"), "external:C:\\Temp\\foo.xlsx#Sheet1!A1",  NULL, "External", "Tip");

    int result = workbook->close(); return result;
}
