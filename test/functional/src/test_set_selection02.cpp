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

    std::shared_ptr<xlsxwriter::workbook> workbook  = std::make_shared<xlsxwriter::workbook>("test_set_selection02.xlsx");
    xlsxwriter::worksheet *worksheet1 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet2 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet3 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet4 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet5 = workbook->add_worksheet();
    xlsxwriter::worksheet *worksheet6 = workbook->add_worksheet();

    worksheet1->set_selection(3, 2, 3, 2);     /* 1. Cell C4. */
    worksheet2->set_selection(3, 2, 6, 6);     /* 2. Cells C4 to G7. */
    worksheet3->set_selection(6, 6, 3, 2);     /* 3. Cells G7 to C.4 */
    worksheet4->set_selection(RANGE("C4:C4")); /* Same as 1. */
    worksheet5->set_selection(RANGE("C4:G7")); /* Same as 2. */
    worksheet6->set_selection(RANGE("G7:C4")); /* Same as 3. */

    int result = workbook->close(); return result;
}
