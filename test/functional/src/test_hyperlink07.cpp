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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_hyperlink07.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_url_opt(CELL("A1"), "external:\\\\VBOXSVR\\share\\foo.xlsx", NULL, "J:\\foo.xlsx", NULL);
    worksheet->write_url(CELL("A3"), "external:foo.xlsx" , NULL);

    int result = workbook->close(); return result;
}
