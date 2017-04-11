/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

#ifndef __LXW_WORKSHEET_HPP__
#define __LXW_WORKSHEET_HPP__

#include <stdio.h>
#include <stdlib.h>
#include <stdint.h>
#include <string.h>

#include <string>
#include <vector>
#include <list>
#include <memory>

#include "xmlwriter.hpp"
#include "shared_strings.hpp"
#include "chart.hpp"
#include "relationships.hpp"
#include "drawing.hpp"
#include "common.hpp"
#include "format.hpp"
#include "utility.hpp"

#define LXW_ROW_MAX 1048576
#define LXW_COL_MAX 16384
#define LXW_COL_META_MAX 128
#define LXW_HEADER_FOOTER_MAX 255
#define LXW_MAX_NUMBER_URLS 65530
#define LXW_PANE_NAME_LENGTH 12 /* bottomRight + 1 */

/* The Excel 2007 specification says that the maximum number of page
 * breaks is 1026. However, in practice it is actually 1023. */
#define LXW_BREAKS_MAX 1023

/** Default column width in Excel */
#define LXW_DEF_COL_WIDTH (double)8.43

/** Default row height in Excel */
#define LXW_DEF_ROW_HEIGHT (double)15.0

namespace xlsxwriter {

/** Gridline options using in `worksheet_gridlines()`. */
enum lxw_gridlines {
    /** Hide screen and print gridlines. */
    LXW_HIDE_ALL_GRIDLINES = 0,
    /** Show screen gridlines. */
    LXW_SHOW_SCREEN_GRIDLINES,
    /** Show print gridlines. */
    LXW_SHOW_PRINT_GRIDLINES,
    /** Show screen and print gridlines. */
    LXW_SHOW_ALL_GRIDLINES
};

enum cell_types {
    NUMBER_CELL = 1,
    STRING_CELL,
    INLINE_STRING_CELL,
    FORMULA_CELL,
    ARRAY_FORMULA_CELL,
    BLANK_CELL,
    BOOLEAN_CELL,
    HYPERLINK_URL,
    HYPERLINK_INTERNAL,
    HYPERLINK_EXTERNAL
};

enum pane_types {
    NO_PANES = 0,
    FREEZE_PANES,
    SPLIT_PANES,
    FREEZE_SPLIT_PANES
};

/* Define the tree.h RB structs for the red-black head types. */
RB_HEAD(lxw_table_cells, lxw_cell);

/* Define a RB_TREE struct manually to add extra members. */
struct lxw_table_rows {
    struct lxw_row *rbh_root;
    struct lxw_row *cached_row;
    lxw_row_t cached_row_num;
};

/* Wrapper around RB_GENERATE_STATIC from tree.h to avoid unused function
 * warnings and to avoid portability issues with the _unused attribute. */
#define LXW_RB_GENERATE_ROW(name, type, field, cmp)       \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)   \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)   \
    RB_GENERATE_INSERT(name, type, field, cmp, static)    \
    RB_GENERATE_REMOVE(name, type, field, static)         \
    RB_GENERATE_FIND(name, type, field, cmp, static)      \
    RB_GENERATE_NEXT(name, type, field, static)           \
    RB_GENERATE_MINMAX(name, type, field, static)         \
    /* Add unused struct to allow adding a semicolon */   \
    struct lxw_rb_generate_row{int unused;}

#define LXW_RB_GENERATE_CELL(name, type, field, cmp)      \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)   \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)   \
    RB_GENERATE_INSERT(name, type, field, cmp, static)    \
    RB_GENERATE_REMOVE(name, type, field, static)         \
    RB_GENERATE_FIND(name, type, field, cmp, static)      \
    RB_GENERATE_NEXT(name, type, field, static)           \
    RB_GENERATE_MINMAX(name, type, field, static)         \
    /* Add unused struct to allow adding a semicolon */   \
    struct lxw_rb_generate_cell{int unused;}

/**
 * @brief Options for rows and columns.
 *
 * Options struct for the worksheet_set_column() and worksheet_set_row()
 * functions.
 *
 * It has the following members but currently only the `hidden` property is
 * supported:
 *
 * * `hidden`
 * * `level`
 * * `collapsed`
 */
struct row_col_options {
    row_col_options() {}
    row_col_options(bool h, uint8_t l, bool c) : hidden(h), level(l), collapsed(c){}
    /** Hide the row/column */
    bool hidden;
    uint8_t level;
    bool collapsed;
};

struct lxw_col_options {
    lxw_col_t firstcol;
    lxw_col_t lastcol;
    double width;
    xlsxwriter::format* format;
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
};

struct lxw_merged_range {
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
};

struct lxw_repeat_rows {
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
};

struct lxw_repeat_cols {
    uint8_t in_use;
    lxw_col_t first_col;
    lxw_col_t last_col;
};

struct lxw_print_area {
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
};

struct lxw_autofilter {
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
};

struct lxw_panes {
    uint8_t type;
    lxw_row_t first_row;
    lxw_col_t first_col;
    lxw_row_t top_row;
    lxw_col_t left_col;
    double x_split;
    double y_split;
};

struct lxw_selection {
    std::string pane;
    std::string active_cell;
    std::string sqref;
};

/**
 * @brief Options for inserted images
 *
 * Options for modifying images inserted via `worksheet_insert_image_opt()`.
 *
 */
struct image_options {

    image_options()
        : x_offset(0)
        , y_offset(0)
        , x_scale(0.0)
        , y_scale(0.0)
        , row(0)
        , col(0)
        , anchor(0)
        , stream(nullptr)
        , image_type(0)
        , width(0.0)
        , height(0.0)
        , x_dpi(0.0)
        , y_dpi(0.0)
        , chart(nullptr)
    {}

    ~image_options(){}
    /** Offset from the left of the cell in pixels. */
    int32_t x_offset;

    /** Offset from the top of the cell in pixels. */
    int32_t y_offset;

    /** X scale of the image as a decimal. */
    double x_scale;

    /** Y scale of the image as a decimal. */
    double y_scale;

    lxw_row_t row;
    lxw_col_t col;
    std::string filename;
    std::string url;
    std::string tip;
    uint8_t anchor;

    /* Internal metadata. */
    FILE *stream;
    uint8_t image_type;
    double width;
    double height;
    std::string short_name;
    std::string extension;
    double x_dpi;
    double y_dpi;
    xlsxwriter::chart* chart;
};

typedef std::shared_ptr<image_options> image_options_ptr;

/**
 * @brief Header and footer options.
 *
 * Optional parameters used in the worksheet_set_header_opt() and
 * worksheet_set_footer_opt() functions.
 *
 */
struct lxw_header_footer_options {
    lxw_header_footer_options() {}
    lxw_header_footer_options(double m) : margin(m) {}
    /** Header or footer margin in inches. Excel default is 0.3. */
    double margin;
};

/**
 * @brief Worksheet protection options.
 */
struct lxw_protection {
    /** Turn off selection of locked cells. This in on in Excel by default.*/
    uint8_t no_select_locked_cells;

    /** Turn off selection of unlocked cells. This in on in Excel by default.*/
    uint8_t no_select_unlocked_cells;

    /** Prevent formatting of cells. */
    uint8_t format_cells;

    /** Prevent formatting of columns. */
    uint8_t format_columns;

    /** Prevent formatting of rows. */
    uint8_t format_rows;

    /** Prevent insertion of columns. */
    uint8_t insert_columns;

    /** Prevent insertion of rows. */
    uint8_t insert_rows;

    /** Prevent insertion of hyperlinks. */
    uint8_t insert_hyperlinks;

    /** Prevent deletion of columns. */
    uint8_t delete_columns;

    /** Prevent deletion of rows. */
    uint8_t delete_rows;

    /** Prevent sorting data. */
    uint8_t sort;

    /** Prevent filtering data. */
    uint8_t autofilter;

    /** Prevent insertion of pivot tables. */
    uint8_t pivot_tables;

    /** Protect scenarios. */
    uint8_t scenarios;

    /** Protect drawing objects. */
    uint8_t objects;

    uint8_t no_sheet;
    uint8_t content;
    uint8_t is_configured;
    char hash[5];
};

class worksheet;

/*
 * Worksheet initialization data.
 */
struct lxw_worksheet_init_data {
    uint32_t index;
    uint8_t hidden;
    uint8_t optimize;
    uint16_t *active_sheet;
    uint16_t *first_sheet;
    sst_ptr sst;
    std::string name;
    std::string quoted_name;
    std::string tmpdir;

};

/* Struct to represent a worksheet row. */
typedef struct lxw_row {
    lxw_row_t row_num;
    double height;
    xlsxwriter::format *format;
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
    uint8_t row_changed;
    uint8_t data_changed;
    uint8_t height_changed;
    lxw_table_cells *cells;

    /* tree management pointers for tree.h. */
    RB_ENTRY (lxw_row) tree_pointers;
} lxw_row;

/* Struct to represent a worksheet cell. */
struct lxw_cell {
    lxw_row_t row_num;
    lxw_col_t col_num;
    enum cell_types type;
    xlsxwriter::format* format;

    union {
        double number;
        int32_t string_id;
        std::string *string;
    } u;

    double formula_result;
    std::string *user_data1;
    std::string *user_data2;
    std::string *sst_string;

    /* List pointers for tree.h. */
    RB_ENTRY (lxw_cell) tree_pointers;
};

class packager;
class workbook;

/**
 * @class worksheet The Worksheet object
 *
 * The Worksheet object represents an Excel worksheet. It handles
 * operations such as writing data to cells or formatting worksheet
 * layout.
 *
 * @brief Functions related to adding data and formatting to a worksheet.
 *
 * The Worksheet object represents an Excel worksheet. It handles
 * operations such as writing data to cells or formatting worksheet
 * layout.
 *
 * A Worksheet object isn't created directly. Instead a worksheet is
 * created by calling the workbook_add_worksheet() function from a
 * Workbook object:
 *
 * @code
 *
 *     #include <xlsxwriter++.hpp>
 *
 *     using namespace xlsxwriter;
 *
 *     int main() {
 *
 *         workbook_ptr workbook  = std::make_shared<workbook>("filename.xlsx");
 *         worksheet_ptr worksheet = workbook->add_worksheet();
 *
 *         worksheet->write_string(0, 0, "Hello Excel", NULL);
 *
 *         return workbook->close();
 *     }
 * @endcode
 *
 */
class worksheet : public xmlwriter{
    friend class xlsxwriter::packager;
    friend class xlsxwriter::workbook;
public:
    worksheet(lxw_worksheet_init_data *init_data);
    ~worksheet();
    /**
     * @brief Write a number to a worksheet cell.
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param number    The number to write to the cell.
     * @param pformat    A pointer to a Format instance or NULL.
     *
     * @return A #lxw_error code.
     *
     * The `write_number()` function writes numeric types to the cell
     * specified by `row` and `column`:
     *
     * @code
     *     worksheet->write_number(0, 0, 123456, NULL);
     *     worksheet->write_number(1, 0, 2.3451, NULL);
     * @endcode
     *
     * @image html write_number01.png
     *
     * The native data type for all numbers in Excel is a IEEE-754 64-bit
     * double-precision floating point, which is also the default type used by
     * `%write_number`.
     *
     * The `format` parameter is used to apply formatting to the cell. This
     * parameter can be `NULL` to indicate no formatting or it can be a
     * @ref format.h "Format" object.
     *
     * @code
     *     xlsxwriter::format* format = workbook.add_format();
     *     format->set_num_format("$#,##0.00");
     *
     *     worksheet->write_number(0, 0, 1234.567, format);
     * @endcode
     *
     * @image html write_number02.png
     *
     */
    lxw_error write_number(lxw_row_t row,
                           lxw_col_t col, double number,
                           format* pformat = nullptr);
    /**
     * @brief Write a string to a worksheet cell.
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param string    String to write to cell.
     * @param pformat    A pointer to a Format instance or NULL.
     *
     * @return A #lxw_error code.
     *
     * The `%write_string()` function writes a string to the cell
     * specified by `row` and `column`:
     *
     * @code
     *     worksheet->write_string(0, 0, "This phrase is English!", NULL);
     * @endcode
     *
     * @image html write_string01.png
     *
     * The `format` parameter is used to apply formatting to the cell. This
     * parameter can be `NULL` to indicate no formatting or it can be a
     * @ref format.h "Format" object:
     *
     * @code
     *     format_ptr format = workbook->add_format();
     *     format->set_bold();
     *
     *     worksheet->write_string(0, 0, "This phrase is Bold!", format);
     * @endcode
     *
     * @image html write_string02.png
     *
     * Unicode strings are supported in UTF-8 encoding. This generally requires
     * that your source file is UTF-8 encoded or that the data has been read from
     * a UTF-8 source:
     *
     * @code
     *    worksheet->write_string(0, 0, "Это фраза на русском!", NULL);
     * @endcode
     *
     * @image html write_string03.png
     *
     */
    lxw_error write_string(lxw_row_t row,
                           lxw_col_t col, const std::string& string,
                           format* pformat = nullptr);
    /**
     * @brief Write a formula to a worksheet cell.
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param formula   Formula string to write to cell.
     * @param format    A pointer to a Format instance or NULL.
     *
     * @return A #lxw_error code.
     *
     * The `%write_formula()` function writes a formula or function to
     * the cell specified by `row` and `column`:
     *
     * @code
     *  worksheet->write_formula(0, 0, "=B3 + 6",                    NULL);
     *  worksheet->write_formula(1, 0, "=SIN(PI()/4)",               NULL);
     *  worksheet->write_formula(2, 0, "=SUM(A1:A2)",                NULL);
     *  worksheet->write_formula(3, 0, "=IF(A3>1,\"Yes\", \"No\")",  NULL);
     *  worksheet->write_formula(4, 0, "=AVERAGE(1, 2, 3, 4)",       NULL);
     *  worksheet->write_formula(5, 0, "=DATEVALUE(\"1-Jan-2013\")", NULL);
     * @endcode
     *
     * @image html write_formula01.png
     *
     * The `format` parameter is used to apply formatting to the cell. This
     * parameter can be `NULL` to indicate no formatting or it can be a
     * @ref format.h "Format" object.
     *
     * Libxlsxwriter doesn't calculate the value of a formula and instead stores a
     * default value of `0`. The correct formula result is displayed in Excel, as
     * shown in the example above, since it recalculates the formulas when it loads
     * the file. For cases where this is an issue see the
     * `worksheet_write_formula_num()` function and the discussion in that section.
     *
     * Formulas must be written with the US style separator/range operator which
     * is a comma (not semi-colon). Therefore a formula with multiple values
     * should be written as follows:
     *
     * @code
     *     // OK.
     *     worksheet->write_formula(0, 0, "=SUM(1, 2, 3)", NULL);
     *
     *     // NO. Error on load.
     *     worksheet->write_formula(1, 0, "=SUM(1; 2; 3)", NULL);
     * @endcode
     *
     */
    lxw_error write_formula(lxw_row_t row,
                            lxw_col_t col, const std::string& formula,
                            format* format);
    /**
     * @brief Write an array formula to a worksheet cell.
     *
     * @param first_row   The first row of the range. (All zero indexed.)
     * @param first_col   The first column of the range.
     * @param last_row    The last row of the range.
     * @param last_col    The last col of the range.
     * @param formula     Array formula to write to cell.
     * @param pformat      A pointer to a Format instance or NULL.
     *
     * @return A #lxw_error code.
     *
      * The `%write_array_formula()` function writes an array formula to
     * a cell range. In Excel an array formula is a formula that performs a
     * calculation on a set of values.
     *
     * In Excel an array formula is indicated by a pair of braces around the
     * formula: `{=SUM(A1:B1*A2:B2)}`.
     *
     * Array formulas can return a single value or a range or values. For array
     * formulas that return a range of values you must specify the range that the
     * return values will be written to. This is why this function has `first_`
     * and `last_` row/column parameters. The RANGE() macro can also be used to
     * specify the range:
     *
     * @code
     *     worksheet->write_array_formula(4, 0, 6, 0,     "{=TREND(C5:C7,B5:B7)}", NULL);
     *
     *     // Same as above using the RANGE() macro.
     *     worksheet->write_array_formula(RANGE("A5:A7"), "{=TREND(C5:C7,B5:B7)}", NULL);
     * @endcode
     *
     * If the array formula returns a single value then the `first_` and `last_`
     * parameters should be the same:
     *
     * @code
     *     worksheet->write_array_formula(1, 0, 1, 0,     "{=SUM(B1:C1*B2:C2)}", NULL);
     *     worksheet->write_array_formula(RANGE("A2:A2"), "{=SUM(B1:C1*B2:C2)}", NULL);
     * @endcode
     *
     */
    lxw_error write_array_formula(lxw_row_t first_row,
                                  lxw_col_t first_col,
                                  lxw_row_t last_row,
                                  lxw_col_t last_col,
                                  const std::string& formula,
                                  format* pformat = nullptr);

    lxw_error write_array_formula_num(lxw_row_t first_row,
                                      lxw_col_t first_col,
                                      lxw_row_t last_row,
                                      lxw_col_t last_col,
                                      const std::string& formula,
                                      format* pformat,
                                      double result);

    /**
     * @brief Write a date or time to a worksheet cell.
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param datetime  The datetime to write to the cell.
     * @param pformat    A pointer to a Format instance or NULL.
     *
     * @return A #lxw_error code.
     *
     * The `worksheet_write_datetime()` function can be used to write a date or
     * time to the cell specified by `row` and `column`:
     *
     * @dontinclude dates_and_times02.c
     * @skip include
     * @until num_format
     * @skip Feb
     * @until }
     *
     * The `format` parameter should be used to apply formatting to the cell using
     * a @ref format.h "Format" object as shown above. Without a date format the
     * datetime will appear as a number only.
     *
     * See @ref working_with_dates for more information about handling dates and
     * times in libxlsxwriter.
     */
    lxw_error write_datetime(lxw_row_t row,
                                       lxw_col_t col, lxw_datetime *datetime,
                                       format* format = nullptr);

    lxw_error write_url_opt(lxw_row_t row_num,
                            lxw_col_t col_num, const std::string& url,
                            format* pformat, const std::string& string = std::string(),
                            const std::string& tooltip = std::string());
    /**
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param url       The url to write to the cell.
     * @param pformat    A pointer to a Format instance or NULL.
     *
     * @return A #lxw_error code.
     *
     *
     * The `%write_url()` function is used to write a URL/hyperlink to a
     * worksheet cell specified by `row` and `column`.
     *
     * @code
     *     worksheet->write_url(0, 0, "http://libxlsxwriter.github.io", url_format);
     * @endcode
     *
     * @image html hyperlinks_short.png
     *
     * The `format` parameter is used to apply formatting to the cell. This
     * parameter can be `NULL` to indicate no formatting or it can be a @ref
     * format.h "Format" object. The typical worksheet format for a hyperlink is a
     * blue underline:
     *
     * @code
     *    format_ptr url_format   = workbook->add_format();
     *
     *    format->set_underline (LXW_UNDERLINE_SINGLE);
     *    format->set_font_color(LXW_COLOR_BLUE);
     *
     * @endcode
     *
     * The usual web style URI's are supported: `%http://`, `%https://`, `%ftp://`
     * and `mailto:` :
     *
     * @code
     *     worksheet->write_url(0, 0, "ftp://www.python.org/",     url_format);
     *     worksheet->write_url(1, 0, "http://www.python.org/",    url_format);
     *     worksheet->write_url(2, 0, "https://www.python.org/",   url_format);
     *     worksheet->write_url(3, 0, "mailto:jmcnamara@cpan.org", url_format);
     *
     * @endcode
     *
     * An Excel hyperlink is comprised of two elements: the displayed string and
     * the non-displayed link. By default the displayed string is the same as the
     * link. However, it is possible to overwrite it with any other
     * `libxlsxwriter` type using the appropriate `write_*()`
     * function. The most common case is to overwrite the displayed link text with
     * another string:
     *
     * @code
     *  // Write a hyperlink but overwrite the displayed string.
     *  worksheet->write_url   (2, 0, "http://libxlsxwriter.github.io", url_format);
     *  worksheet->write_string(2, 0, "Read the documentation.",        url_format);
     *
     * @endcode
     *
     * @image html hyperlinks_short2.png
     *
     * Two local URIs are supported: `internal:` and `external:`. These are used
     * for hyperlinks to internal worksheet references or external workbook and
     * worksheet references:
     *
     * @code
     *     worksheet->write_url(0, 0, "internal:Sheet2!A1",                url_format);
     *     worksheet->write_url(1, 0, "internal:Sheet2!B2",                url_format);
     *     worksheet->write_url(2, 0, "internal:Sheet2!A1:B2",             url_format);
     *     worksheet->write_url(3, 0, "internal:'Sales Data'!A1",          url_format);
     *     worksheet->write_url(4, 0, "external:c:\\temp\\foo.xlsx",       url_format);
     *     worksheet->write_url(5, 0, "external:c:\\foo.xlsx#Sheet2!A1",   url_format);
     *     worksheet->write_url(6, 0, "external:..\\foo.xlsx",             url_format);
     *     worksheet->write_url(7, 0, "external:..\\foo.xlsx#Sheet2!A1",   url_format);
     *     worksheet->write_url(8, 0, "external:\\\\NET\\share\\foo.xlsx", url_format);
     *
     * @endcode
     *
     * Worksheet references are typically of the form `Sheet1!A1`. You can also
     * link to a worksheet range using the standard Excel notation:
     * `Sheet1!A1:B2`.
     *
     * In external links the workbook and worksheet name must be separated by the
     * `#` character:
     *
     * @code
     *     worksheet->write_url(0, 0, "external:c:\\foo.xlsx#Sheet2!A1",   url_format);
     * @endcode
     *
     * You can also link to a named range in the target worksheet: For example say
     * you have a named range called `my_name` in the workbook `c:\temp\foo.xlsx`
     * you could link to it as follows:
     *
     * @code
     *     worksheet->write_url(0, 0, "external:c:\\temp\\foo.xlsx#my_name", url_format);
     *
     * @endcode
     *
     * Excel requires that worksheet names containing spaces or non alphanumeric
     * characters are single quoted as follows:
     *
     * @code
     *     worksheet->write_url(0, 0, "internal:'Sales Data'!A1", url_format);
     * @endcode
     *
     * Links to network files are also supported. Network files normally begin
     * with two back slashes as follows `\\NETWORK\etc`. In order to represent
     * this in a C string literal the backslashes should be escaped:
     * @code
     *     worksheet->write_url(0, 0, "external:\\\\NET\\share\\foo.xlsx", url_format);
     * @endcode
     *
     *
     * Alternatively, you can use Windows style forward slashes. These are
     * translated internally to backslashes:
     *
     * @code
     *     worksheet->write_url(0, 0, "external:c:/temp/foo.xlsx",     url_format);
     *     worksheet->write_url(1, 0, "external://NET/share/foo.xlsx", url_format);
     *
     * @endcode
     *
     *
     * **Note:**
     *
     *    libxlsxwriter will escape the following characters in URLs as required
     *    by Excel: `\s " < > \ [ ]  ^ { }` unless the URL already contains `%%xx`
     *    style escapes. In which case it is assumed that the URL was escaped
     *    correctly by the user and will by passed directly to Excel.
     *
     */
    lxw_error write_url(lxw_row_t row,
                        lxw_col_t col, const std::string& url,
                        format* pformat = nullptr);

    /**
     * @brief Write a formatted boolean worksheet cell.
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param value     The boolean value to write to the cell.
     * @param pformat    A pointer to a Format instance or NULL.
     *
     * @return A #lxw_error code.
     *
     * Write an Excel boolean to the cell specified by `row` and `column`:
     *
     * @code
     *     worksheet->write_boolean(2, 2, 0, my_format);
     * @endcode
     *
     */
    lxw_error write_boolean(lxw_row_t row, lxw_col_t col, bool value, format* pformat = nullptr);

    /**
     * @brief Write a formatted blank worksheet cell.
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param format    A pointer to a Format instance or NULL.
     *
     * @return A #lxw_error code.
     *
     * Write a blank cell specified by `row` and `column`:
     *
     * @code
     *     worksheet->write_blank(1, 1, border_format);
     * @endcode
     *
     * This function is used to add formatting to a cell which doesn't contain a
     * string or number value.
     *
     * Excel differentiates between an "Empty" cell and a "Blank" cell. An Empty
     * cell is a cell which doesn't contain data or formatting whilst a Blank cell
     * doesn't contain data but does contain formatting. Excel stores Blank cells
     * but ignores Empty cells.
     *
     * As such, if you write an empty cell without formatting it is ignored.
     *
     */
    lxw_error write_blank(lxw_row_t row, lxw_col_t col, format* pformat = nullptr);

    /**
     * @brief Write a formula to a worksheet cell with a user defined result.
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param formula   Formula string to write to cell.
     * @param pformat    A pointer to a Format instance or NULL.
     * @param result    A user defined result for a formula.
     *
     * @return A #lxw_error code.
     *
     * The `%write_formula_num()` function writes a formula or Excel
     * function to the cell specified by `row` and `column` with a user defined
     * result:
     *
     * @code
     *     // Required as a workaround only.
     *     worksheet->write_formula_num(0, 0, "=1 + 2", NULL, 3);
     * @endcode
     *
     * Libxlsxwriter doesn't calculate the value of a formula and instead stores
     * the value `0` as the formula result. It then sets a global flag in the XLSX
     * file to say that all formulas and functions should be recalculated when the
     * file is opened.
     *
     * This is the method recommended in the Excel documentation and in general it
     * works fine with spreadsheet applications.
     *
     * However, applications that don't have a facility to calculate formulas,
     * such as Excel Viewer, or some mobile applications will only display the `0`
     * results.
     *
     * If required, the `%write_formula_num()` function can be used to
     * specify a formula and its result.
     *
     * This function is rarely required and is only provided for compatibility
     * with some third party applications. For most applications the
     * worksheet->write_formula() function is the recommended way of writing
     * formulas.
     *
     */
    lxw_error write_formula_num(lxw_row_t row,
                                lxw_col_t col,
                                const std::string& formula,
                                format* pformat, double result);

    /**
     * @brief Set the properties for a row of cells.
     *
     * @param row       The zero indexed row number.
     * @param height    The row height.
     * @param format    A pointer to a Format instance or NULL.
     *
     * The `%set_row()` function is used to change the default
     * properties of a row. The most common use for this function is to change the
     * height of a row:
     *
     * @code
     *     // Set the height of Row 1 to 20.
     *     worksheet->set_row(0, 20, NULL);
     * @endcode
     *
     * The other common use for `%worksheet_set_row()` is to set the a @ref
     * format.h "Format" for all cells in the row:
     *
     * @code
     *     xlsxwriter::format *bold = workbook->add_format();
     *     format->set_bold(bold);
     *
     *     // Set the header row to bold.
     *     worksheet->set_row(0, 15, bold);
     * @endcode
     *
     * If you wish to set the format of a row without changing the height you can
     * pass the default row height of #LXW_DEF_ROW_HEIGHT = 15:
     *
     * @code
     *     worksheet->set_row(0, LXW_DEF_ROW_HEIGHT, format);
     *     worksheet->set_row(0, 15, format); // Same as above.
     * @endcode
     *
     * The `format` parameter will be applied to any cells in the row that don't
     * have a format. As with Excel the row format is overridden by an explicit
     * cell format. For example:
     *
     * @code
     *     // Row 1 has format1.
     *     worksheet->set_row(0, 15, format1);
     *
     *     // Cell A1 in Row 1 defaults to format1.
     *     worksheet->write_string(0, 0, "Hello", NULL);
     *
     *     // Cell B1 in Row 1 keeps format2.
     *     worksheet->write_string(0, 1, "Hello", format2);
     * @endcode
     *
     */
    lxw_error set_row(lxw_row_t row, double height, format* format = nullptr);

    /**
     * @brief Set the properties for a row of cells.
     *
     * @param row       The zero indexed row number.
     * @param height    The row height.
     * @param pformat    A pointer to a Format instance or NULL.
     * @param options   Optional row parameters: hidden, level, collapsed.
     *
     * The `%set_row_opt()` function  is the same as
     *  `set_row()` with an additional `options` parameter.
     *
     * The `options` parameter is a #lxw_row_col_options struct. It has the
     * following members but currently only the `hidden` property is supported:
     *
     * - `hidden`
     * - `level`
     * - `collapsed`
     *
     * The `"hidden"` option is used to hide a row. This can be used, for
     * example, to hide intermediary steps in a complicated calculation:
     *
     * @code
     *     row_col_options options = {};
     *     options.hidden = 1;
     *     options.level = 0;
     *     options.collapsed = 0;
     *
     *     // Hide the fourth row.
     *     worksheet->set_row(3, 20, NULL, options);
     * @endcode
     *
     */
    lxw_error set_row_opt(lxw_row_t row,
                          double height,
                          format* pformat,
                          const row_col_options& options = row_col_options(false, 0, false));

    /**
     * @brief Set the properties for one or more columns of cells.
     *
     * @param first_col The zero indexed first column.
     * @param last_col  The zero indexed last column.
     * @param width     The width of the column(s).
     * @param pformat    A pointer to a Format instance or NULL.
     *
     * The `%set_column()` function can be used to change the default
     * properties of a single column or a range of columns:
     *
     * @code
     *     // Width of columns B:D set to 30.
     *     worksheet->set_column(1, 3, 30, NULL);
     *
     * @endcode
     *
     * If `%set_column()` is applied to a single column the value of
     * `first_col` and `last_col` should be the same:
     *
     * @code
     *     // Width of column B set to 30.
     *     worksheet->set_column(1, 1, 30, NULL);
     *
     * @endcode
     *
     * It is also possible, and generally clearer, to specify a column range using
     * the form of `COLS()` macro:
     *
     * @code
     *     worksheet->set_column(4, 4, 20, NULL);
     *     worksheet->set_column(5, 8, 30, NULL);
     *
     *     // Same as the examples above but clearer.
     *     worksheet->set_column(COLS("E:E"), 20, NULL);
     *     worksheet->set_column(COLS("F:H"), 30, NULL);
     *
     * @endcode
     *
     * The `width` parameter sets the column width in the same units used by Excel
     * which is: the number of characters in the default font. The default width
     * is 8.43 in the default font of Calibri 11. The actual relationship between
     * a string width and a column width in Excel is complex. See the
     * [following explanation of column widths](https://support.microsoft.com/en-us/kb/214123)
     * from the Microsoft support documentation for more details.
     *
     * There is no way to specify "AutoFit" for a column in the Excel file
     * format. This feature is only available at runtime from within Excel. It is
     * possible to simulate "AutoFit" in your application by tracking the maximum
     * width of the data in the column as your write it and then adjusting the
     * column width at the end.
     *
     * As usual the @ref format.h `format` parameter is optional. If you wish to
     * set the format without changing the width you can pass a default column
     * width of #LXW_DEF_COL_WIDTH = 8.43:
     *
     * @code
     *     format* format = workbook->add_format();
     *     format->set_bold();
     *
     *     // Set the first column to bold.
     *     worksheet->set_column(0, 0, LXW_DEF_COL_HEIGHT, format);
     * @endcode
     *
     * The `format` parameter will be applied to any cells in the column that
     * don't have a format. For example:
     *
     * @code
     *     // Column 1 has format1.
     *     worksheet->set_column(COLS("A:A"), 8.43, format1);
     *
     *     // Cell A1 in column 1 defaults to format1.
     *     worksheet->write_string(0, 0, "Hello", NULL);
     *
     *     // Cell A2 in column 1 keeps format2.
     *     worksheet->write_string(1, 0, "Hello", format2);
     * @endcode
     *
     * As in Excel a row format takes precedence over a default column format:
     *
     * @code
     *     // Row 1 has format1.
     *     worksheet->set_row(0, 15, format1);
     *
     *     // Col 1 has format2.
     *     worksheet->set_column(COLS("A:A"), 8.43, format2);
     *
     *     // Cell A1 defaults to format1, the row format.
     *     worksheet->write_string(0, 0, "Hello", NULL);
     *
     *    // Cell A2 keeps format2, the column format.
     *     worksheet->write_string(1, 0, "Hello", NULL);
     * @endcode
     */
    lxw_error set_column(lxw_col_t first_col,
                         lxw_col_t last_col,
                         double width, format* format = nullptr);

     /**
      * @brief Set the properties for one or more columns of cells with options.
      *
      * @param first_col The zero indexed first column.
      * @param last_col  The zero indexed last column.
      * @param width     The width of the column(s).
      * @param pformat    A pointer to a Format instance or NULL.
      * @param options   Optional row parameters: hidden, level, collapsed.
      *
      * The `%set_column_opt()` function  is the same as
      * `worksheet_set_column()` with an additional `options` parameter.
      *
      * The `options` parameter is a #row_col_options struct. It has the
      * following members but currently only the `hidden` property is supported:
      *
      * - `hidden`
      * - `level`
      * - `collapsed`
      *
      * The `"hidden"` option is used to hide a column. This can be used, for
      * example, to hide intermediary steps in a complicated calculation:
      *
      * @code
      *     row_col_options options = {};
      *     options.hidden = 1;
      *     options.level = 0;
      *     options.collapsed = 0;
      *
      *     worksheet->set_column_opt(COLS("A:A"), 8.43, NULL, options);
      * @endcode
      *
      */
    lxw_error set_column_opt(lxw_col_t first_col,
                             lxw_col_t last_col,
                             double width,
                             format* pformat,
                             const row_col_options& options = row_col_options(false, 0, false));

    /**
     * @brief Insert an image in a worksheet cell.
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param filename  The image filename, with path if required.
     *
     * @return A #lxw_error code.
     *
     * This function can be used to insert a image into a worksheet. The image can
     * be in PNG, JPEG or BMP format:
     *
     * @code
     *     worksheet->insert_image(2, 1, "logo.png");
     * @endcode
     *
     * @image html insert_image.png
     *
     * The `insert_image_opt()` function takes additional optional
     * parameters to position and scale the image, see below.
     *
     * **Note**:
     * The scaling of a image may be affected if is crosses a row that has its
     * default height changed due to a font that is larger than the default font
     * size or that has text wrapping turned on. To avoid this you should
     * explicitly set the height of the row using `set_row()` if it
     * crosses an inserted image.
     *
     * BMP images are only supported for backward compatibility. In general it is
     * best to avoid BMP images since they aren't compressed. If used, BMP images
     * must be 24 bit, true color, bitmaps.
     */
    lxw_error insert_image(lxw_row_t row, lxw_col_t col,
                           const std::string& filename);

    /**
     * @brief Insert an image in a worksheet cell, with options.
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param filename  The image filename, with path if required.
     * @param options   Optional image parameters.
     *
     * @return A #lxw_error code.
     *
     * The `%insert_image_opt()` function is like
     * `insert_image()` function except that it takes an optional
     * #image_options struct to scale and position the image:
     *
     * @code
     *    image_options options = {}
     *    options.x_offset = 30;
     *    options.y_offset = 10;
     *    options.x_scale  = 0.5;
     *    options.y_scale  = 0.5;
     *
     *    worksheet->insert_image_opt(2, 1, "logo.png", options);
     *
     * @endcode
     *
     * @image html insert_image_opt.png
     *
     * @note See the notes about row scaling and BMP images in
     * `insert_image()` above.
     */
    lxw_error insert_image_opt(lxw_row_t row, lxw_col_t col,
                               const std::string& filename,
                               image_options* options = nullptr);
    /**
     * @brief Insert a chart object into a worksheet.
     *
     * @param row       The zero indexed row number.
     * @param col       The zero indexed column number.
     * @param chart     A #chart object created via workbook->add_chart().
     *
     * @return A #lxw_error code.
     *
     * The `%insert_chart()` can be used to insert a chart into a
     * worksheet. The chart object must be created first using the
     * `add_chart()` function and configured using the @ref chart.hpp
     * functions.
     *
     * @code
     *     // Create a chart object.
     *     chart_ptr chart = workbook->add_chart(LXW_CHART_LINE);
     *
     *     // Add a data series to the chart.
     *     chart->add_series("", "=Sheet1!$A$1:$A$6");
     *
     *     // Insert the chart into the worksheet
     *     worksheet->insert_chart(0, 2, chart);
     * @endcode
     *
     * @image html chart_working.png
     *
     *
     * **Note:**
     *
     * A chart may only be inserted into a worksheet once. If several similar
     * charts are required then each one must be created separately with
     * `%insert_chart()`.
     *
     */
    lxw_error insert_chart(lxw_row_t row, lxw_col_t col, xlsxwriter::chart* chart);

    /**
     * @brief Insert a chart object into a worksheet, with options.
     *
     * @param row          The zero indexed row number.
     * @param col          The zero indexed column number.
     * @param chart        A #xlsxwriter::chart object created via workbook->add_chart().
     * @param user_options Optional chart parameters.
     *
     * @return A #lxw_error code.
     *
     * The `%insert_chart_opt()` function is like
     * `insert_chart()` function except that it takes an optional
     * #image_options struct to scale and position the image of the chart:
     *
     * @code
     *    image_options options = {}
     *    options.x_offset = 30;
     *    options.y_offset = 10;
     *    options.x_scale  = 0.5;
     *    options.y_scale  = 0.75;
     *
     *    worksheet->insert_chart_opt(0, 2, chart, options);
     *
     * @endcode
     *
     * @image html chart_line_opt.png
     *
     * The #image_options struct is the same struct used in
     * `insert_image_opt()` to position and scale images.
     *
     */
    lxw_error insert_chart_opt(lxw_row_t row, lxw_col_t col,
                               xlsxwriter::chart* chart,
                               image_options* user_options);

    /**
     * @brief Merge a range of cells.
     *
     * @param first_row The first row of the range. (All zero indexed.)
     * @param first_col The first column of the range.
     * @param last_row  The last row of the range.
     * @param last_col  The last col of the range.
     * @param string    String to write to the merged range.
     * @param pformat    A pointer to a Format instance or NULL.
     *
     * @return A #lxw_error code.
     *
     * The `%merge_range()` function allows cells to be merged together
     * so that they act as a single area.
     *
     * Excel generally merges and centers cells at same time. To get similar
     * behavior with libxlsxwriter you need to apply a @ref format.hpp "Format"
     * object with the appropriate alignment:
     *
     * @code
     *     format_ptr merge_format = workbook->add_format();
     *     merge_format->set_align(LXW_ALIGN_CENTER);
     *
     *     worksheet->merge_range(1, 1, 1, 3, "Merged Range", merge_format);
     *
     * @endcode
     *
     * It is possible to apply other formatting to the merged cells as well:
     *
     * @code
     *    merge_format->set_align   (LXW_ALIGN_CENTER);
     *    merge_format->set_align   (LXW_ALIGN_VERTICAL_CENTER);
     *    merge_format->set_border  (LXW_BORDER_DOUBLE);
     *    merge_format->set_bold    ();
     *    merge_format->set_bg_color(0xD7E4BC);
     *
     *    worksheet->merge_range(2, 1, 3, 3, "Merged Range", merge_format);
     *
     * @endcode
     *
     * @image html merge.png
     *
     * The `%merge_range()` function writes a `char*` string using
     * `write_string()`. In order to write other data types, such as a
     * number or a formula, you can overwrite the first cell with a call to one of
     * the other write functions. The same Format should be used as was used in
     * the merged range.
     *
     * @code
     *    // First write a range with a blank string.
     *    worksheet->merge_range (1, 1, 1, 3, "", format);
     *
     *    // Then overwrite the first cell with a number.
     *    worksheet->write_number(1, 1, 123, format);
     * @endcode
     *
     * @note Merged ranges generally don’t work in libxlsxwriter when the Workbook
     * #xlsxwriter::workbook_options `constant_memory` mode is enabled.
     */
    lxw_error merge_range(lxw_row_t first_row,
                          lxw_col_t first_col, lxw_row_t last_row,
                          lxw_col_t last_col, const std::string& string,
                          format* pformat);

    /**
     * @brief Set the autofilter area in the worksheet.
     *
     * @param first_row The first row of the range. (All zero indexed.)
     * @param first_col The first column of the range.
     * @param last_row  The last row of the range.
     * @param last_col  The last col of the range.
     *
     * @return A #lxw_error code.
     *
     * The `%autofilter()` function allows an autofilter to be added to
     * a worksheet.
     *
     * An autofilter is a way of adding drop down lists to the headers of a 2D
     * range of worksheet data. This allows users to filter the data based on
     * simple criteria so that some data is shown and some is hidden.
     *
     * @image html autofilter.png
     *
     * To add an autofilter to a worksheet:
     *
     * @code
     *     worksheet->autofilter(0, 0, 50, 3);
     *
     *     // Same as above using the RANGE() macro.
     *     worksheet->autofilter(RANGE("A1:D51"));
     * @endcode
     *
     * Note: it isn't currently possible to apply filter conditions to the
     * autofilter.
     */
    lxw_error autofilter(lxw_row_t first_row,
                         lxw_col_t first_col, lxw_row_t last_row,
                         lxw_col_t last_col);

     /**
      * @brief Make a worksheet the active, i.e., visible worksheet.
      *
      * The `%activate()` function is used to specify which worksheet is
      * initially visible in a multi-sheet workbook:
      *
      * @code
      *     worksheet_ptr worksheet1 = workbook->add_worksheet();
      *     worksheet_ptr worksheet2 = workbook->add_worksheet();
      *     worksheet_ptr worksheet3 = workbook->add_worksheet();
      *
      *     worksheet3->activate();
      * @endcode
      *
      * @image html worksheet_activate.png
      *
      * More than one worksheet can be selected via the `worksheet_select()`
      * function, see below, however only one worksheet can be active.
      *
      * The default active worksheet is the first worksheet.
      *
      */
    void activate();

     /**
      * @brief Set a worksheet tab as selected.
      *
      * The `%select()` function is used to indicate that a worksheet is
      * selected in a multi-sheet workbook:
      *
      * @code
      *     worksheet1->activate();
      *     worksheet2->select();
      *     worksheet3->select();
      *
      * @endcode
      *
      * A selected worksheet has its tab highlighted. Selecting worksheets is a
      * way of grouping them together so that, for example, several worksheets
      * could be printed in one go. A worksheet that has been activated via the
      * `activate()` function will also appear as selected.
      *
      */
    void select();

    /**
     * @brief Hide the current worksheet.
     *
     * The `%hide()` function is used to hide a worksheet:
     *
     * @code
     *     worksheet2->hide();
     * @endcode
     *
     * You may wish to hide a worksheet in order to avoid confusing a user with
     * intermediate data or calculations.
     *
     * @image html hide_sheet.png
     *
     * A hidden worksheet can not be activated or selected so this function is
     * mutually exclusive with the `worksheet_activate()` and `worksheet_select()`
     * functions. In addition, since the first worksheet will default to being the
     * active worksheet, you cannot hide the first worksheet without activating
     * another sheet:
     *
     * @code
     *     worksheet1->activate();
     *     worksheet2->hide();
     * @endcode
     */
    void hide();

    /**
     * @brief Set current worksheet as the first visible sheet tab.
     *
     * The `activate()` function determines which worksheet is initially
     * selected.  However, if there are a large number of worksheets the selected
     * worksheet may not appear on the screen. To avoid this you can select the
     * leftmost visible worksheet tab using `%worksheet_set_first_sheet()`:
     *
     * @code
     *     worksheet19->set_first_sheet(); // First visible worksheet tab.
     *     worksheet20->activate();        // First visible worksheet.
     * @endcode
     *
     * This function is not required very often. The default value is the first
     * worksheet.
     */
    void set_first_sheet();

    /**
     * @brief Split and freeze a worksheet into panes.
     *
     * @param row       The cell row (zero indexed).
     * @param col       The cell column (zero indexed).
     *
     * The `%worksheet_freeze_panes()` function can be used to divide a worksheet
     * into horizontal or vertical regions known as panes and to "freeze" these
     * panes so that the splitter bars are not visible.
     *
     * The parameters `row` and `col` are used to specify the location of the
     * split. It should be noted that the split is specified at the top or left of
     * a cell and that the function uses zero based indexing. Therefore to freeze
     * the first row of a worksheet it is necessary to specify the split at row 2
     * (which is 1 as the zero-based index).
     *
     * You can set one of the `row` and `col` parameters as zero if you do not
     * want either a vertical or horizontal split.
     *
     * Examples:
     *
     * @code
     *     worksheet1->freeze_panes(1, 0); // Freeze the first row.
     *     worksheet2->freeze_panes(0, 1); // Freeze the first column.
     *     worksheet3->freeze_panes(1, 1); // Freeze first row/column.
     *
     * @endcode
     *
     */
    void freeze_panes(lxw_row_t row, lxw_col_t col);
    /**
     * @brief Split a worksheet into panes.
     *
     * @param vertical   The position for the vertical split.
     * @param horizontal The position for the horizontal split.
     *
     * The `%split_panes()` function can be used to divide a worksheet
     * into horizontal or vertical regions known as panes. This function is
     * different from the `worksheet_freeze_panes()` function in that the splits
     * between the panes will be visible to the user and each pane will have its
     * own scroll bars.
     *
     * The parameters `vertical` and `horizontal` are used to specify the vertical
     * and horizontal position of the split. The units for `vertical` and
     * `horizontal` are the same as those used by Excel to specify row height and
     * column width. However, the vertical and horizontal units are different from
     * each other. Therefore you must specify the `vertical` and `horizontal`
     * parameters in terms of the row heights and column widths that you have set
     * or the default values which are 15 for a row and 8.43 for a column.
     *
     * Examples:
     *
     * @code
     *     worksheet1->split_panes(15, 0);    // First row.
     *     worksheet2->split_panes(0,  8.43); // First column.
     *     worksheet3->split_panes(15, 8.43); // First row and column.
     *
     * @endcode
     *
     */
    void split_panes(double vertical, double horizontal);

    /* freeze_panes() with infrequent options. Undocumented for now. */
    void freeze_panes_opt(lxw_row_t first_row, lxw_col_t first_col,
                          lxw_row_t top_row, lxw_col_t left_col,
                          uint8_t type);

    /* split_panes() with infrequent options. Undocumented for now. */
    void split_panes_opt(double vertical, double horizontal,
                         lxw_row_t top_row, lxw_col_t left_col);
    /**
     * @brief Set the selected cell or cells in a worksheet:
     *
     * @param first_row   The first row of the range. (All zero indexed.)
     * @param first_col   The first column of the range.
     * @param last_row    The last row of the range.
     * @param last_col    The last col of the range.
     *
     *
     * The `%set_selection()` function can be used to specify which cell
     * or range of cells is selected in a worksheet: The most common requirement
     * is to select a single cell, in which case the `first_` and `last_`
     * parameters should be the same.
     *
     * The active cell within a selected range is determined by the order in which
     * `first_` and `last_` are specified.
     *
     * Examples:
     *
     * @code
     *     worksheet1->set_selection(3, 3, 3, 3);     // Cell D4.
     *     worksheet2->set_selection(3, 3, 6, 6);     // Cells D4 to G7.
     *     worksheet3->set_selection(6, 6, 3, 3);     // Cells G7 to D4.
     *     worksheet5->set_selection(RANGE("D4:G7")); // Using the RANGE macro.
     *
     * @endcode
     *
     */
    void set_selection(lxw_row_t first_row, lxw_col_t first_col,
                       lxw_row_t last_row, lxw_col_t last_col);

    /**
     * @brief Set the page orientation as landscape.
     *
     * @param worksheet Pointer to a lxw_worksheet instance to be updated.
     *
     * This function is used to set the orientation of a worksheet's printed page
     * to landscape:
     *
     * @code
     *     worksheet->set_landscape();
     * @endcode
     */
    void set_landscape();

    /**
     * @brief Set the page orientation as portrait.
     *
     * This function is used to set the orientation of a worksheet's printed page
     * to portrait. The default worksheet orientation is portrait, so this
     * function isn't generally required:
     *
     * @code
     *     worksheet.set_portrait();
     * @endcode
     */
    void set_portrait();

    /**
     * @brief Set the page layout to page view mode.
     *
     * This function is used to display the worksheet in "Page View/Layout" mode:
     *
     * @code
     *     worksheet.set_page_view();
     * @endcode
     */
    void set_page_view();

    /**
     * @brief Set the paper type for printing.
     *
     * @param paper_type The Excel paper format type.
     *
     * This function is used to set the paper format for the printed output of a
     * worksheet. The following paper styles are available:
     *
     *
     *   Index    | Paper format            | Paper size
     *   :------- | :---------------------- | :-------------------
     *   0        | Printer default         | Printer default
     *   1        | Letter                  | 8 1/2 x 11 in
     *   2        | Letter Small            | 8 1/2 x 11 in
     *   3        | Tabloid                 | 11 x 17 in
     *   4        | Ledger                  | 17 x 11 in
     *   5        | Legal                   | 8 1/2 x 14 in
     *   6        | Statement               | 5 1/2 x 8 1/2 in
     *   7        | Executive               | 7 1/4 x 10 1/2 in
     *   8        | A3                      | 297 x 420 mm
     *   9        | A4                      | 210 x 297 mm
     *   10       | A4 Small                | 210 x 297 mm
     *   11       | A5                      | 148 x 210 mm
     *   12       | B4                      | 250 x 354 mm
     *   13       | B5                      | 182 x 257 mm
     *   14       | Folio                   | 8 1/2 x 13 in
     *   15       | Quarto                  | 215 x 275 mm
     *   16       | ---                     | 10x14 in
     *   17       | ---                     | 11x17 in
     *   18       | Note                    | 8 1/2 x 11 in
     *   19       | Envelope 9              | 3 7/8 x 8 7/8
     *   20       | Envelope 10             | 4 1/8 x 9 1/2
     *   21       | Envelope 11             | 4 1/2 x 10 3/8
     *   22       | Envelope 12             | 4 3/4 x 11
     *   23       | Envelope 14             | 5 x 11 1/2
     *   24       | C size sheet            | ---
     *   25       | D size sheet            | ---
     *   26       | E size sheet            | ---
     *   27       | Envelope DL             | 110 x 220 mm
     *   28       | Envelope C3             | 324 x 458 mm
     *   29       | Envelope C4             | 229 x 324 mm
     *   30       | Envelope C5             | 162 x 229 mm
     *   31       | Envelope C6             | 114 x 162 mm
     *   32       | Envelope C65            | 114 x 229 mm
     *   33       | Envelope B4             | 250 x 353 mm
     *   34       | Envelope B5             | 176 x 250 mm
     *   35       | Envelope B6             | 176 x 125 mm
     *   36       | Envelope                | 110 x 230 mm
     *   37       | Monarch                 | 3.875 x 7.5 in
     *   38       | Envelope                | 3 5/8 x 6 1/2 in
     *   39       | Fanfold                 | 14 7/8 x 11 in
     *   40       | German Std Fanfold      | 8 1/2 x 12 in
     *   41       | German Legal Fanfold    | 8 1/2 x 13 in
     *
     * Note, it is likely that not all of these paper types will be available to
     * the end user since it will depend on the paper formats that the user's
     * printer supports. Therefore, it is best to stick to standard paper types:
     *
     * @code
     *     worksheet->set_paper(1);  // US Letter
     *     worksheet->set_paper(9);  // A4
     * @endcode
     *
     * If you do not specify a paper type the worksheet will print using the
     * printer's default paper style.
     */
    void set_paper(uint8_t paper_type);

    /**
     * @brief Set the worksheet margins for the printed page.
     *
     * @param left    Left margin in inches.   Excel default is 0.7.
     * @param right   Right margin in inches.  Excel default is 0.7.
     * @param top     Top margin in inches.    Excel default is 0.75.
     * @param bottom  Bottom margin in inches. Excel default is 0.75.
     *
     * The `%worksheet_set_margins()` function is used to set the margins of the
     * worksheet when it is printed. The units are in inches. Specifying `-1` for
     * any parameter will give the default Excel value as shown above.
     *
     * @code
     *    worksheet->set_margins(1.3, 1.2, -1, -1);
     * @endcode
     *
     */
    void set_margins(double left, double right, double top, double bottom);

    /**
     * @brief Set the printed page header caption.
     *
     * @param string    The header string.
     *
     * @return A #lxw_error code.
     *
     * Headers and footers are generated using a string which is a combination of
     * plain text and control characters.
     *
     * The available control character are:
     *
     *
     *   | Control         | Category      | Description           |
     *   | --------------- | ------------- | --------------------- |
     *   | `&L`            | Justification | Left                  |
     *   | `&C`            |               | Center                |
     *   | `&R`            |               | Right                 |
     *   | `&P`            | Information   | Page number           |
     *   | `&N`            |               | Total number of pages |
     *   | `&D`            |               | Date                  |
     *   | `&T`            |               | Time                  |
     *   | `&F`            |               | File name             |
     *   | `&A`            |               | Worksheet name        |
     *   | `&Z`            |               | Workbook path         |
     *   | `&fontsize`     | Font          | Font size             |
     *   | `&"font,style"` |               | Font name and style   |
     *   | `&U`            |               | Single underline      |
     *   | `&E`            |               | Double underline      |
     *   | `&S`            |               | Strikethrough         |
     *   | `&X`            |               | Superscript           |
     *   | `&Y`            |               | Subscript             |
     *
     *
     * Text in headers and footers can be justified (aligned) to the left, center
     * and right by prefixing the text with the control characters `&L`, `&C` and
     * `&R`.
     *
     * For example (with ASCII art representation of the results):
     *
     * @code
     *     worksheet.set_header("&LHello");
     *
     *      ---------------------------------------------------------------
     *     |                                                               |
     *     | Hello                                                         |
     *     |                                                               |
     *
     *
     *     worksheet.set_header("&CHello");
     *
     *      ---------------------------------------------------------------
     *     |                                                               |
     *     |                          Hello                                |
     *     |                                                               |
     *
     *
     *     worksheet.set_header("&RHello");
     *
     *      ---------------------------------------------------------------
     *     |                                                               |
     *     |                                                         Hello |
     *     |                                                               |
     *
     *
     * @endcode
     *
     * For simple text, if you do not specify any justification the text will be
     * centered. However, you must prefix the text with `&C` if you specify a font
     * name or any other formatting:
     *
     * @code
     *     worksheet.set_header("Hello");
     *
     *      ---------------------------------------------------------------
     *     |                                                               |
     *     |                          Hello                                |
     *     |                                                               |
     *
     * @endcode
     *
     * You can have text in each of the justification regions:
     *
     * @code
     *     worksheet->set_header("&LCiao&CBello&RCielo");
     *
     *      ---------------------------------------------------------------
     *     |                                                               |
     *     | Ciao                     Bello                          Cielo |
     *     |                                                               |
     *
     * @endcode
     *
     * The information control characters act as variables that Excel will update
     * as the workbook or worksheet changes. Times and dates are in the users
     * default format:
     *
     * @code
     *     worksheet->set_header("&CPage &P of &N");
     *
     *      ---------------------------------------------------------------
     *     |                                                               |
     *     |                        Page 1 of 6                            |
     *     |                                                               |
     *
     *     worksheet->set_header("&CUpdated at &T");
     *
     *      ---------------------------------------------------------------
     *     |                                                               |
     *     |                    Updated at 12:30 PM                        |
     *     |                                                               |
     *
     * @endcode
     *
     * You can specify the font size of a section of the text by prefixing it with
     * the control character `&n` where `n` is the font size:
     *
     * @code
     *     worksheet->set_header("&C&30Hello Big");
     *     worksheet->set_header("&C&10Hello Small");
     *
     * @endcode
     *
     * You can specify the font of a section of the text by prefixing it with the
     * control sequence `&"font,style"` where `fontname` is a font name such as
     * Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
     * "Courier New" or "Times New Roman" and `style` is one of the standard
     *
     * @code
     *     worksheet->set_header("&C&\"Courier New,Italic\"Hello");
     *     worksheet->set_header("&C&\"Courier New,Bold Italic\"Hello");
     *     worksheet->set_header("&C&\"Times New Roman,Regular\"Hello");
     *
     * @endcode
     *
     * It is possible to combine all of these features together to create
     * sophisticated headers and footers. As an aid to setting up complicated
     * headers and footers you can record a page set-up as a macro in Excel and
     * look at the format strings that VBA produces. Remember however that VBA
     * uses two double quotes `""` to indicate a single double quote. For the last
     * example above the equivalent VBA code looks like this:
     *
     * @code
     *     .LeftHeader = ""
     *     .CenterHeader = "&""Times New Roman,Regular""Hello"
     *     .RightHeader = ""
     *
     * @endcode
     *
     * Alternatively you can inspect the header and footer strings in an Excel
     * file by unzipping it and grepping the XML sub-files. The following shows
     * how to do that using libxml's xmllint to format the XML for clarity:
     *
     * @code
     *
     *    $ unzip myfile.xlsm -d myfile
     *    $ xmllint --format `find myfile -name "*.xml" | xargs` | egrep "Header|Footer"
     *
     *      <headerFooter scaleWithDoc="0">
     *        <oddHeader>&amp;L&amp;P</oddHeader>
     *      </headerFooter>
     *
     * @endcode
     *
     * Note that in this case you need to unescape the Html. In the above example
     * the header string would be `&L&P`.
     *
     * To include a single literal ampersand `&` in a header or footer you should
     * use a double ampersand `&&`:
     *
     * @code
     *     worksheet->set_header("&CCuriouser && Curiouser - Attorneys at Law");
     * @endcode
     *
     * Note, the header or footer string must be less than 255 characters. Strings
     * longer than this will not be written.
     *
     */
    lxw_error set_header(const std::string& header_string);

    /**
     * @brief Set the printed page footer caption.
     *
     * @param footer_string    The footer string.
     *
     * @return A #lxw_error code.
     *
     * The syntax of this function is the same as set_header().
     *
     */
    lxw_error set_footer(const std::string& footer_string);

    /**
     * @brief Set the printed page header caption with additional options.
     *
     * @param header_string    The header string.
     * @param options          Header options.
     *
     * @return A #lxw_error code.
     *
     * The syntax of this function is the same as worksheet_set_header() with an
     * additional parameter to specify options for the header.
     *
     * Currently, the only available option is the header margin:
     *
     * @code
     *
     *    lxw_header_footer_options header_options = { 0.2 };
     *
     *    worksheet->set_header_opt("Some text", &header_options);
     *
     * @endcode
     *
     */
    lxw_error set_header_opt(const std::string& header_string, const lxw_header_footer_options& options = lxw_header_footer_options(0));

    /**
     * @brief Set the printed page footer caption with additional options.
     *
     * @param footer_string    The footer string.
     * @param options          Footer options.
     *
     * @return A #lxw_error code.
     *
     * The syntax of this function is the same as set_header_opt().
     *
     */
    lxw_error set_footer_opt(const std::string& footer_string, const lxw_header_footer_options& options = lxw_header_footer_options(0));

    /**
     * @brief Set the horizontal page breaks on a worksheet.
     *
     * @param breaks    Array of page breaks.
     *
     * @return A #lxw_error code.
     *
     * The `%set_h_pagebreaks()` function adds horizontal page breaks to
     * a worksheet. A page break causes all the data that follows it to be printed
     * on the next page. Horizontal page breaks act between rows.
     *
     * The function takes an array of one or more page breaks. The type of the
     * array data is @ref lxw_row_t and the last element of the array must be 0:
     *
     * @code
     *    std::vector<lxw_row_t> breaks1 = {20, 0}; // 1 page break. Zero indicates the end.
     *    std::vector<lxw_row_t> breaks2 = {20, 40, 60, 80, 0};
     *
     *    worksheet->set_h_pagebreaks(breaks1);
     *    worksheet->set_h_pagebreaks(breaks2);
     * @endcode
     *
     * To create a page break between rows 20 and 21 you must specify the break at
     * row 21. However in zero index notation this is actually row 20:
     *
     * @code
     *    // Break between row 20 and 21.
     *    std::vector<lxw_row_t> breaks = {20, 0};
     *
     *    worksheet->set_h_pagebreaks(breaks);
     * @endcode
     *
     * There is an Excel limitation of 1023 horizontal page breaks per worksheet.
     *
     * Note: If you specify the "fit to page" option via the
     * `fit_to_pages()` function it will override all manual page
     * breaks.
     *
     */
    lxw_error set_h_pagebreaks(const std::vector<lxw_row_t>& breaks);

    /**
     * @brief Set the vertical page breaks on a worksheet.
     *
     * @param breaks    Array of page breaks.
     *
     * @return A #lxw_error code.
     *
     * The `%set_v_pagebreaks()` function adds vertical page breaks to a
     * worksheet. A page break causes all the data that follows it to be printed
     * on the next page. Vertical page breaks act between columns.
     *
     * The function takes an array of one or more page breaks. The type of the
     * array data is @ref lxw_col_t and the last element of the array must be 0:
     *
     * @code
     *    std::vector<lxw_col_t> breaks1 = {20, 0}; // 1 page break. Zero indicates the end.
     *    std::vector<lxw_col_t> breaks2 = {20, 40, 60, 80, 0};
     *
     *    worksheet->set_v_pagebreaks(breaks1);
     *    worksheet->set_v_pagebreaks(breaks2);
     * @endcode
     *
     * To create a page break between columns 20 and 21 you must specify the break
     * at column 21. However in zero index notation this is actually column 20:
     *
     * @code
     *    // Break between column 20 and 21.
     *    std::vector<lxw_col_t> breaks = {20, 0};
     *
     *    worksheet->set_v_pagebreaks(breaks);
     * @endcode
     *
     * There is an Excel limitation of 1023 vertical page breaks per worksheet.
     *
     * Note: If you specify the "fit to page" option via the
     * `worksheet_fit_to_pages()` function it will override all manual page
     * breaks.
     *
     */
    lxw_error set_v_pagebreaks(const std::vector<lxw_col_t>& breaks);

    /**
     * @brief Set the order in which pages are printed.
     *
     * The `%print_across()` function is used to change the default
     * print direction. This is referred to by Excel as the sheet "page order":
     *
     * @code
     *     worksheet->print_across();
     * @endcode
     *
     * The default page order is shown below for a worksheet that extends over 4
     * pages. The order is called "down then across":
     *
     *     [1] [3]
     *     [2] [4]
     *
     * However, by using the `print_across` function the print order will be
     * changed to "across then down":
     *
     *     [1] [2]
     *     [3] [4]
     *
     */
    void print_across();

    /**
     * @brief Set the worksheet zoom factor.
     *
     * @param scale     Worksheet zoom factor.
     *
     * Set the worksheet zoom factor in the range `10 <= zoom <= 400`:
     *
     * @code
     *     worksheet->set_zoom(50);
     *     worksheet->set_zoom(75);
     *     worksheet->set_zoom(300);
     *     worksheet->set_zoom(400);
     * @endcode
     *
     * The default zoom factor is 100. It isn't possible to set the zoom to
     * "Selection" because it is calculated by Excel at run-time.
     *
     * Note, `%zoom()` does not affect the scale of the printed
     * page. For that you should use `set_print_scale()`.
     */
    void set_zoom(uint16_t scale);

    /**
     * @brief Set the option to display or hide gridlines on the screen and
     *        the printed page.
     *
     * @param option    Gridline option.
     *
     * Display or hide screen and print gridlines using one of the values of
     * @ref lxw_gridlines.
     *
     * @code
     *    worksheet->gridlines(LXW_HIDE_ALL_GRIDLINES);
     *
     *    worksheet->gridlines(LXW_SHOW_PRINT_GRIDLINES);
     * @endcode
     *
     * The Excel default is that the screen gridlines are on  and the printed
     * worksheet is off.
     *
     */
    void gridlines(uint8_t option);

    /**
     * @brief Center the printed page horizontally.
     *
     * Center the worksheet data horizontally between the margins on the printed
     * page:
     *
     * @code
     *     worksheet->center_horizontally();
     * @endcode
     *
     */
    void center_horizontally();

    /**
     * @brief Center the printed page vertically.
     *
     * Center the worksheet data vertically between the margins on the printed
     * page:
     *
     * @code
     *     worksheet->center_vertically();
     * @endcode
     *
     */
    void center_vertically();

    /**
     * @brief Set the option to print the row and column headers on the printed
     *        page.
     *
     * When printing a worksheet from Excel the row and column headers (the row
     * numbers on the left and the column letters at the top) aren't printed by
     * default.
     *
     * This function sets the printer option to print these headers:
     *
     * @code
     *    worksheet->print_row_col_headers();
     * @endcode
     *
     */
    void print_row_col_headers();

    /**
     * @brief Set the number of rows to repeat at the top of each printed page.
     *
     * @param first_row First row of repeat range.
     * @param last_row  Last row of repeat range.
     *
     * @return A #lxw_error code.
     *
     * For large Excel documents it is often desirable to have the first row or
     * rows of the worksheet print out at the top of each page.
     *
     * This can be achieved by using this function. The parameters `first_row`
     * and `last_row` are zero based:
     *
     * @code
     *     worksheet->repeat_rows(0, 0); // Repeat the first row.
     *     worksheet->repeat_rows(0, 1); // Repeat the first two rows.
     * @endcode
     */
    lxw_error repeat_rows(lxw_row_t first_row, lxw_row_t last_row);

    /**
     * @brief Set the number of columns to repeat at the top of each printed page.
     *
     * @param first_col First column of repeat range.
     * @param last_col  Last column of repeat range.
     *
     * @return A #lxw_error code.
     *
     * For large Excel documents it is often desirable to have the first column or
     * columns of the worksheet print out at the left of each page.
     *
     * This can be achieved by using this function. The parameters `first_col`
     * and `last_col` are zero based:
     *
     * @code
     *     worksheet->repeat_columns(0, 0); // Repeat the first col.
     *     worksheet->repeat_columns(0, 1); // Repeat the first two cols.
     * @endcode
     */
    lxw_error repeat_columns(lxw_col_t first_col, lxw_col_t last_col);

    /**
     * @brief Set the print area for a worksheet.
     *
     * @param worksheet Pointer to a lxw_worksheet instance to be updated.
     * @param first_row The first row of the range. (All zero indexed.)
     * @param first_col The first column of the range.
     * @param last_row  The last row of the range.
     * @param last_col  The last col of the range.
     *
     * @return A #lxw_error code.
     *
     * This function is used to specify the area of the worksheet that will be
     * printed. The RANGE() macro is often convenient for this.
     *
     * @code
     *     worksheet->print_area(0, 0, 41, 10); // A1:K42.
     *
     *     // Same as:
     *     worksheet->print_area(RANGE("A1:K42"));
     * @endcode
     *
     * In order to set a row or column range you must specify the entire range:
     *
     * @code
     *     worksheet->print_area(RANGE("A1:H1048576")); // Same as A:H.
     * @endcode
     */
    lxw_error print_area(lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
    /**
     * @brief Fit the printed area to a specific number of pages both vertically
     *        and horizontally.
     *
     * @param width     Number of pages horizontally.
     * @param height    Number of pages vertically.
     *
     * The `%fit_to_pages()` function is used to fit the printed area to
     * a specific number of pages both vertically and horizontally. If the printed
     * area exceeds the specified number of pages it will be scaled down to
     * fit. This ensures that the printed area will always appear on the specified
     * number of pages even if the page size or margins change:
     *
     * @code
     *     worksheet->fit_to_pages(1, 1); // Fit to 1x1 pages.
     *     worksheet->fit_to_pages(2, 1); // Fit to 2x1 pages.
     *     worksheet->fit_to_pages(1, 2); // Fit to 1x2 pages.
     * @endcode
     *
     * The print area can be defined using the `print_area()` function
     * as described above.
     *
     * A common requirement is to fit the printed output to `n` pages wide but
     * have the height be as long as necessary. To achieve this set the `height`
     * to zero:
     *
     * @code
     *     // 1 page wide and as long as necessary.
     *     worksheet->fit_to_pages(worksheet, 1, 0);
     * @endcode
     *
     * **Note**:
     *
     * - Although it is valid to use both `%fit_to_pages()` and
     *   `worksheet_set_print_scale()` on the same worksheet Excel only allows one
     *   of these options to be active at a time. The last function call made will
     *   set the active option.
     *
     * - The `%fit_to_pages()` function will override any manual page
     *   breaks that are defined in the worksheet.
     *
     * - When using `%fit_to_pages()` it may also be required to set the
     *   printer paper size using `set_paper()` or else Excel will
     *   default to "US Letter".
     *
     */
    void fit_to_pages(uint16_t width, uint16_t height);

    /**
     * @brief Set the start page number when printing.
     *
     * @param start_page Starting page number.
     *
     * The `%set_start_page()` function is used to set the number of
     * the starting page when the worksheet is printed out:
     *
     * @code
     *     // Start print from page 2.
     *     worksheet->set_start_page(2);
     * @endcode
     */
    void set_start_page(uint16_t start_page);

    /**
     * @brief Set the scale factor for the printed page.
     *
     * @param scale     Print scale of worksheet to be printed.
     *
     * This function sets the scale factor of the printed page. The Scale factor
     * must be in the range `10 <= scale <= 400`:
     *
     * @code
     *     worksheet->set_print_scale(75);
     *     worksheet->set_print_scale(400);
     * @endcode
     *
     * The default scale factor is 100. Note, `%worksheet_set_print_scale()` does
     * not affect the scale of the visible page in Excel. For that you should use
     * `worksheet_set_zoom()`.
     *
     * Note that although it is valid to use both `worksheet_fit_to_pages()` and
     * `%worksheet_set_print_scale()` on the same worksheet Excel only allows one
     * of these options to be active at a time. The last function call made will
     * set the active option.
     *
     */
    void set_print_scale(uint16_t scale);

    /**
     * @brief Display the worksheet cells from right to left for some versions of
     *        Excel.
     *
     * The `%right_to_left()` function is used to change the default
     * direction of the worksheet from left-to-right, with the `A1` cell in the
     * top left, to right-to-left, with the `A1` cell in the top right.
     *
     * @code
     *     worksheet->get_right_to_left();
     * @endcode
     *
     * This is useful when creating Arabic, Hebrew or other near or far eastern
     * worksheets that use right-to-left as the default direction.
     */
    void get_right_to_left();

    /**
     * @brief Hide zero values in worksheet cells.
     *
     * The `%hide_zero()` function is used to hide any zero values that
     * appear in cells:
     *
     * @code
     *     worksheet->hide_zero();
     * @endcode
     */
    void hide_zero();

    /**
     * @brief Set the color of the worksheet tab.
     *
     * @param color     The tab color.
     *
     * The `%set_tab_color()` function is used to change the color of the worksheet
     * tab:
     *
     * @code
     *      worksheet_set_tab_color(xlsxwriter::LXW_COLOR_RED);
     *      worksheet_set_tab_color(LXW_COLOR_GREEN);
     *      worksheet_set_tab_color(0xFF9900); // Orange.
     * @endcode
     *
     * The color should be an RGB integer value, see @ref working_with_colors.
     */
    void set_tab_color(lxw_color_t color);

    /**
     * @brief Protect elements of a worksheet from modification.
     *
     * @param password  A worksheet password.
     * @param options   Worksheet elements to protect.
     *
     * The `%protect()` function protects worksheet elements from modification:
     *
     * @code
     *     worksheet->protect("Some Password", options);
     * @endcode
     *
     * The `password` and lxw_protection pointer are both optional:
     *
     * @code
     *     worksheet_protect(NULL,       NULL);
     *     worksheet_protect(NULL,       my_options);
     *     worksheet_protect("password", NULL);
     *     worksheet_protect("password", my_options);
     * @endcode
     *
     * Passing a `NULL` password is the same as turning on protection without a
     * password. Passing a `NULL` password and `NULL` options, or any other
     * combination has the effect of enabling a cell's `locked` and `hidden`
     * properties if they have been set.
     *
     * A *locked* cell cannot be edited and this property is on by default for all
     * cells. A *hidden* cell will display the results of a formula but not the
     * formula itself. These properties can be set using the format_set_unlocked()
     * and format_set_hidden() format functions.
     *
     * You can specify which worksheet elements you wish to protect by passing a
     * lxw_protection pointer in the `options` argument with any or all of the
     * following members set:
     *
     *     no_select_locked_cells
     *     no_select_unlocked_cells
     *     format_cells
     *     format_columns
     *     format_rows
     *     insert_columns
     *     insert_rows
     *     insert_hyperlinks
     *     delete_columns
     *     delete_rows
     *     sort
     *     autofilter
     *     pivot_tables
     *     scenarios
     *     objects
     *
     * All parameters are off by default. Individual elements can be protected as
     * follows:
     *
     * @code
     *     lxw_protection options = {
     *         .format_cells             = 1,
     *         .insert_hyperlinks        = 1,
     *         .insert_rows              = 1,
     *         .delete_rows              = 1,
     *         .insert_columns           = 1,
     *         .delete_columns           = 1,
     *     };
     *
     *     worksheet->protect(NULL, &options);
     *
     * @endcode
     *
     * See also the format_set_unlocked() and format_set_hidden() format functions.
     *
     * **Note:** Worksheet level passwords in Excel offer **very** weak
     * protection. They don't encrypt your data and are very easy to
     * deactivate. Full workbook encryption is not supported by `libxlsxwriter`
     * since it requires a completely different file format and would take several
     * man months to implement.
     */
    void protect(const char *password, lxw_protection *options);

    /**
     * @brief Set the default row properties.
     *
     * @param worksheet        Pointer to a lxw_worksheet instance to be updated.
     * @param height           Default row height.
     * @param hide_unused_rows Hide unused cells.
     *
     * The `%set_default_row()` function is used to set Excel default
     * row properties such as the default height and the option to hide unused
     * rows. These parameters are an optimization used by Excel to set row
     * properties without generating a very large file with an entry for each row.
     *
     * To set the default row height:
     *
     * @code
     *     worksheet_set_default_row(24, false);
     *
     * @endcode
     *
     * To hide unused rows:
     *
     * @code
     *     worksheet_set_default_row(15, true);
     * @endcode
     *
     * Note, in the previous case we use the default height #LXW_DEF_ROW_HEIGHT =
     * 15 so the the height remains unchanged.
     */
    void set_default_row(double height, uint8_t hide_unused_rows);

    void set_vertical_dpi(size_t dpi);

    void assemble_xml_file();
    void write_single_row();

    void prepare_image(uint16_t image_ref_id, uint16_t drawing_id,
                       const image_options_ptr& image_data);

    void prepare_chart(uint16_t chart_ref_id, uint16_t drawing_id,
                       const image_options_ptr& image_data);

    lxw_row *find_row(lxw_row_t row_num);
    lxw_cell *find_cell(lxw_row *row, lxw_col_t col_num);

private:
    FILE *optimize_tmpfile;
    lxw_table_rows *table;
    lxw_table_rows *hyperlinks;
    lxw_cell **array;
    std::vector<std::shared_ptr<lxw_merged_range>> merged_ranges;
    std::vector<std::shared_ptr<lxw_selection>> selections;
    std::vector<std::shared_ptr<image_options>> image_data;
    std::vector<std::shared_ptr<image_options>> chart_data;

    lxw_row_t dim_rowmin;
    lxw_row_t dim_rowmax;
    lxw_col_t dim_colmin;
    lxw_col_t dim_colmax;

    sst_ptr sst;
    std::string name;
    std::string quoted_name;
    std::string tmpdir;

    uint32_t index;
    uint8_t active;
    uint8_t selected;
    uint8_t hidden;
    uint16_t *active_sheet;
    uint16_t *first_sheet;

    lxw_col_options **col_options;
    uint16_t col_options_max;

    double *col_sizes;
    uint16_t col_sizes_max;

    xlsxwriter::format **col_formats;
    uint16_t col_formats_max;

    uint8_t col_size_changed;
    uint8_t row_size_changed;
    uint8_t optimize;
    lxw_row *optimize_row;

    uint16_t fit_height;
    uint16_t fit_width;
    uint16_t horizontal_dpi;
    uint16_t hlink_count;
    uint16_t page_start;
    uint16_t print_scale;
    uint16_t rel_count;
    uint16_t vertical_dpi;
    uint16_t zoom;
    uint8_t filter_on;
    uint8_t fit_page;
    uint8_t hcenter;
    uint8_t orientation;
    uint8_t outline_changed;
    uint8_t outline_on;
    uint8_t page_order;
    uint8_t page_setup_changed;
    uint8_t page_view;
    uint8_t paper_size;
    uint8_t print_gridlines;
    uint8_t print_headers;
    uint8_t print_options_changed;
    uint8_t right_to_left;
    uint8_t screen_gridlines;
    uint8_t show_zeros;
    uint8_t vba_codename;
    uint8_t vcenter;
    uint8_t zoom_scale_normal;

    lxw_color_t tab_color;

    double margin_left;
    double margin_right;
    double margin_top;
    double margin_bottom;
    double margin_header;
    double margin_footer;

    double default_row_height;
    uint32_t default_row_pixels;
    uint32_t default_col_pixels;
    uint8_t default_row_zeroed;
    uint8_t default_row_set;

    uint8_t header_footer_changed;
    std::string header;
    std::string footer;

    lxw_repeat_rows repeat_rows_;
    lxw_repeat_cols repeat_cols_;
    lxw_print_area print_area_;
    lxw_autofilter autofilter_;

    uint16_t merged_range_count;

    std::vector<lxw_row_t> hbreaks;
    std::vector<lxw_col_t> vbreaks;
    uint16_t hbreaks_count;
    uint16_t vbreaks_count;

    std::list<rel_tuple_ptr> external_hyperlinks;
    std::list<rel_tuple_ptr> external_drawing_links;
    std::list<rel_tuple_ptr> drawing_links;

    lxw_panes panes;

    lxw_protection protection;

    std::shared_ptr<xlsxwriter::drawing> drawing;

    /* Declarations required for unit testing. */
    void _xml_declaration();
    void _write_worksheet();
    void _write_dimension();
    void _write_sheet_view();
    void _write_sheet_views();
    void _write_sheet_format_pr();
    void _write_sheet_data();
    void _write_page_margins();
    void _write_page_setup();
    void _write_col_info(lxw_col_options *options);
    void _write_row(lxw_row *row, const std::string& spans);

    void _write_merge_cell(const std::shared_ptr<lxw_merged_range>& merged_range);
    void _write_merge_cells();

    void _write_odd_header();
    void _write_odd_footer();
    void _write_header_footer();

    void _write_print_options();
    void _write_sheet_pr();
    void _write_tab_color();
    void _write_sheet_protection();
    void _write_optimized_sheet_data();
    uint32_t _calculate_x_split_width(double x_split) const;
    void _write_cell(lxw_cell *cell, xlsxwriter::format* row_format);
    void _write_rows();
    void _write_drawing(uint16_t id);
    void _write_drawings();
    void _position_object_emus(const image_options_ptr &image, const drawing_object_ptr &drawing_object);
    void _write_formula_num_cell(lxw_cell *cell);
    void _write_array_formula_num_cell(lxw_cell *cell);
    void _write_inline_string_cell(const std::string &range, int32_t style_index, lxw_cell *cell);
    void _write_freeze_panes();
    void _position_object_pixels(const image_options_ptr &image, const drawing_object_ptr &drawing_object);
    void _write_string_cell(const std::string &range, int32_t style_index, lxw_cell *cell);
    void _write_number_cell(const std::string &range, int32_t style_index, lxw_cell *cell);
    void _write_split_panes();
    void _write_selection(const std::shared_ptr<lxw_selection> &selection);
    int32_t _size_col(lxw_col_t col_num);
    void _write_hyperlink_external(lxw_row_t row_num, lxw_col_t col_num, const std::string* location, const std::string *tooltip, uint16_t id);
    void _write_hyperlinks();
    void _write_hyperlink_internal(lxw_row_t row_num, lxw_col_t col_num, const std::string* location, const std::string *display, const std::string *tooltip);
    void _write_page_set_up_pr();
    void _write_cols();
    int32_t _size_row(lxw_row_t row_num);
    void _write_panes();
    void _write_selections();
    lxw_row *_get_row(lxw_row_t row_num);
    lxw_error _check_dimensions(lxw_row_t row_num, lxw_col_t col_num, int8_t ignore_row, int8_t ignore_col);
    void _write_brk(uint32_t id, uint32_t max);
    void _write_row_breaks();
    void _write_auto_filter();
    void _write_col_breaks();
    void _write_boolean_cell(lxw_cell *cell);
    void _insert_cell(lxw_row_t row_num, lxw_col_t col_num, lxw_cell *cell);
    void _insert_hyperlink(lxw_row_t row_num, lxw_col_t col_num, lxw_cell *link);
};

lxw_row *_get_row_list(lxw_table_rows *table, lxw_row_t row_num);
typedef std::shared_ptr<worksheet> worksheet_ptr;

} // xlsxwriter

#endif /* __LXW_WORKSHEET_H__ */
