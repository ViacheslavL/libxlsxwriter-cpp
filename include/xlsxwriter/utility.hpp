/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @file utility.h
 *
 * @brief Utility functions for libxlsxwriter.
 *
 * <!-- Copyright 2014-2016, John McNamara, jmcnamara@cpan.org -->
 *
 */

#ifndef __LXW_UTILITY_H__
#define __LXW_UTILITY_H__

#include <stdint.h>
#include "common.hpp"
#include <string>

namespace xlsxwriter {

/**
 * @brief Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *      worksheet->write_string(CELL("A1"), "Foo", NULL);
 *
 *      //Same as:
 *      worksheet->write_string(0, 0,       "Foo", NULL);
 * @endcode
 *
 * @note
 *
 * This macro shouldn't be used in performance critical situations since it
 * expands to two function calls.
 */
#define CELL(cell) \
    xlsxwriter::lxw_name_to_row(cell), xlsxwriter::lxw_name_to_col(cell)

/**
 * @brief Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *     worksheet->set_column(COLS("B:D"), 20, NULL, NULL);
 *
 *     // Same as:
 *     worksheet->set_column(1, 3,        20, NULL, NULL);
 * @endcode
 *
 */
#define COLS(cols) \
    xlsxwriter::lxw_name_to_col(cols), xlsxwriter::lxw_name_to_col_2(cols)

/**
 * @brief Convert an Excel `A1:B2` range into a `(first_row, first_col,
 *        last_row, last_col)` sequence.
 *
 * Convert an Excel `A1:B2` range into a `(first_row, first_col, last_row,
 * last_col)` sequence.
 *
 * This is a little syntactic shortcut to help with worksheet layout.
 *
 * @code
 *     worksheet->print_area(0, 0, 41, 10); // A1:K42.
 *
 *     // Same as:
 *     worksheet->print_area(RANGE("A1:K42"));
 * @endcode
 */
#define RANGE(range) \
    xlsxwriter::lxw_name_to_row(range), xlsxwriter::lxw_name_to_col(range), \
    xlsxwriter::lxw_name_to_row_2(range), xlsxwriter::lxw_name_to_col_2(range)


/**
 * @brief Converts a libxlsxwriter error number to a string.
 *
 * The `%lxw_strerror` function converts a libxlsxwriter error number defined
 * by #lxw_error to a pointer to a string description of the error.
 * Similar to the standard library strerror(3) function.
 *
 * For example:
 *
 * @code
 *     lxw_error error = workbook->close();
 *
 *     if (error)
 *         printf("Error in workbook_close().\n"
 *                "Error %d = %s\n", error, lxw_strerror(error));
 * @endcode
 *
 * This would produce output like the following if the target file wasn't
 * writable:
 *
 *     Error in workbook_close().
 *     Error 2 = Error creating output xlsx file. Usually a permissions error.
 *
 * @param error_num The error number returned by a libxlsxwriter function.
 *
 * @return A pointer to a statically allocated string. Do not free.
 */
const std::string& lxw_strerror(lxw_error error_num);

/* Create a quoted version of the worksheet name */
std::string lxw_quote_sheetname(const std::string& str);

void lxw_col_to_name(std::string& col_name, lxw_col_t col_num, uint8_t absolute);

void lxw_rowcol_to_cell(std::string& cell_name, lxw_row_t row, lxw_col_t col);

void lxw_rowcol_to_cell_abs(std::string& cell_name,
                            lxw_row_t row,
                            lxw_col_t col, uint8_t abs_row, uint8_t abs_col);

void lxw_rowcol_to_range(std::string& range,
                         lxw_row_t first_row, lxw_col_t first_col,
                         lxw_row_t last_row, lxw_col_t last_col);

void lxw_rowcol_to_range_abs(std::string& range,
                             lxw_row_t first_row, lxw_col_t first_col,
                             lxw_row_t last_row, lxw_col_t last_col);

void lxw_rowcol_to_formula_abs(std::string& formula, const std::string& sheetname,
                          lxw_row_t first_row, lxw_col_t first_col,
                          lxw_row_t last_row, lxw_col_t last_col);

uint32_t lxw_name_to_row(const std::string& row_str);
uint16_t lxw_name_to_col(const std::string&  col_str);
uint32_t lxw_name_to_row_2(const std::string& row_str);
uint16_t lxw_name_to_col_2(const std::string& col_str);

double lxw_datetime_to_excel_date(lxw_datetime *datetime, uint8_t date_1904);

void lxw_str_tolower(std::string& str);

FILE *lxw_tmpfile(const char *tmpdir);

std::string to_string(double);

} // namespace xlsxwriter

#endif /* __LXW_UTILITY_H__ */
