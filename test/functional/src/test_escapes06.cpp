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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_escapes06.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    xlsxwriter::format *num_format = workbook->add_format();

    num_format->set_num_format("[Red]0.0%\\ \"a\"");

    worksheet->set_column(0, 0, 14, NULL);

    worksheet->write_number(CELL("A1"), 123, num_format);

    int result = workbook->close(); return result;
}
