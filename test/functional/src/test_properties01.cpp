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

    xlsxwriter::workbook       *workbook   = new xlsxwriter::workbook("test_properties01.xlsx");
    xlsxwriter::worksheet      *worksheet  = workbook->add_worksheet();
    xlsxwriter::doc_properties properties = {};

    properties.title    = strdup("This is an example spreadsheet");
    properties.subject  = strdup("With document properties");
    properties.author   = strdup("Someone");
    properties.manager  = strdup("Dr. Heinz Doofenshmirtz");
    properties.company  = strdup("of Wolves");
    properties.category = strdup("Example spreadsheets");
    properties.keywords = strdup("Sample, Example, Properties");
    properties.comments = strdup("Created with Perl and Excel::Writer::XLSX");
    properties.status   = strdup("Quo");

    workbook->set_properties(properties);

    worksheet->set_column(0, 0, 70);
    worksheet->write_string(CELL("A1"), "Select 'Office Button -> Prepare -> Properties' to see the file properties.");

    int result = workbook->close(); return result;
}
