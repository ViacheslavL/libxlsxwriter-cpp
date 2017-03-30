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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_gh42_02.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    std::string string = "\0xe5\0x9b\0xbe\0x20\0xe5\0x9b\0xbe\0x00";

    worksheet->write_string(0, 0, string, NULL);

    workbook->close();

    return 0;
}

