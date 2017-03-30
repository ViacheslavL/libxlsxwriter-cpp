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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_array_formula01.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();

    worksheet->write_number(0, 1, 0, NULL);
    worksheet->write_number(1, 1, 0, NULL);
    worksheet->write_number(2, 1, 0, NULL);
    worksheet->write_number(0, 2, 0, NULL);
    worksheet->write_number(1, 2, 0, NULL);
    worksheet->write_number(2, 2, 0, NULL);

    worksheet->write_array_formula(RANGE("A1:A3"), "{=SUM(B1:C1*B2:C2)}");

    int result = workbook->close(); return result;
}
