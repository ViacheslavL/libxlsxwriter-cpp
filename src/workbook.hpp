/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @page workbook_page The Workbook object
 *
 * The Workbook is the main object exposed by the libxlsxwriter library. It
 * represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
 *
 * See @ref workbook.hpp for full details of the functionality.
 *
 * @file workbook.hpp
 *
 * @brief Functions related to creating an Excel xlsx workbook.
 *
 * The Workbook is the main object exposed by the libxlsxwriter library. It
 * represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
 *
 * @code
 *     #include "xlsxwriter.h"
 *
 *     int main() {
 *
 *         workbook_ptr  workbook  = std::make_shared<workbook>("filename.xlsx");
 *         worksheet_ptr worksheet = workbook->add_worksheet();
 *
 *         worksheet->write_string(0, 0, "Hello Excel");
 *
 *         return workbook->close();
 *     }
 * @endcode
 *
 * @image html workbook01.png
 *
 */
#ifndef __LXW_WORKBOOK_HPP__
#define __LXW_WORKBOOK_HPP__

#include <stdint.h>
#include <stdio.h>
#include <errno.h>

#include "worksheet.hpp"
#include "chart.hpp"
#include "shared_strings.hpp"
#include "hash_table.hpp"
#include "common.hpp"

#include <map>
#include <unordered_set>

#define LXW_DEFINED_NAME_LENGTH 128

/* Define the tree.h RB structs for the red-black head types. */
RB_HEAD(lxw_worksheet_names, lxw_worksheet_name);

/* Define the queue.h structs for the workbook lists. */
STAILQ_HEAD(lxw_worksheets, lxw_worksheet);
STAILQ_HEAD(lxw_charts, lxw_chart);
TAILQ_HEAD(lxw_defined_names, lxw_defined_name);

namespace xlsxwriter {

/* Struct to represent a worksheet name/pointer pair. */
typedef struct lxw_worksheet_name {
    const char *name;
    lxw_worksheet *worksheet;

    RB_ENTRY (lxw_worksheet_name) tree_pointers;
} lxw_worksheet_name;

/* Wrapper around RB_GENERATE_STATIC from tree.h to avoid unused function
 * warnings and to avoid portability issues with the _unused attribute. */
#define LXW_RB_GENERATE_NAMES(name, type, field, cmp)     \
    RB_GENERATE_INSERT_COLOR(name, type, field, static)   \
    RB_GENERATE_REMOVE_COLOR(name, type, field, static)   \
    RB_GENERATE_INSERT(name, type, field, cmp, static)    \
    RB_GENERATE_REMOVE(name, type, field, static)         \
    RB_GENERATE_FIND(name, type, field, cmp, static)      \
    RB_GENERATE_NEXT(name, type, field, static)           \
    RB_GENERATE_MINMAX(name, type, field, static)         \
    /* Add unused struct to allow adding a semicolon */   \
    struct lxw_rb_generate_names{int unused;}

/* Struct to represent a defined name. */
struct defined_name {
    int16_t index;
    bool hidden;
    std::string name;
    std::string app_name;
    std::string formula;
    std::string normalised_name;
    std::string normalised_sheetname;

    /* List pointers for queue.h. */
    TAILQ_ENTRY (lxw_defined_name) list_pointers;
};

typedef std::shared_ptr<defined_name> defined_name_ptr;

/**
 * Workbook document properties.
 */
struct doc_properties {
    /** The title of the Excel Document. */
    std::string title;

    /** The subject of the Excel Document. */
    std::string subject;

    /** The author of the Excel Document. */
    std::string author;

    /** The manager field of the Excel Document. */
    std::string manager;

    /** The company field of the Excel Document. */
    std::string company;

    /** The category of the Excel Document. */
    std::string category;

    /** The keywords of the Excel Document. */
    std::string keywords;

    /** The comment field of the Excel Document. */
    std::string comments;

    /** The status of the Excel Document. */
    std::string status;

    /** The hyperlink base url of the Excel Document. */
    std::string hyperlink_base;

    time_t created;

};

/**
 * @brief Workbook options.
 *
 * Optional parameters when creating a new Workbook object via
 * workbook_new_opt().
 *
 * The following properties are supported:
 *
 * - `constant_memory`: Reduces the amount of data stored in memory so that
 *   large files can be written efficiently.
 *
 *   @note In this mode a row of data is written and then discarded when a
 *   cell in a new row is added via one of the `worksheet->write_*()`
 *   methods. Therefore, once this option is active, data should be written in
 *   sequential row order. For this reason the `worksheet->merge_range()`
 *   doesn't work in this mode. See also @ref ww_mem_constant.
 *
 * - `tmpdir`: libxlsxwriter stores workbook data in temporary files prior
 *   to assembling the final XLSX file. The temporary files are created in the
 *   system's temp directory. If the default temporary directory isn't
 *   accessible to your application, or doesn't contain enough space, you can
 *   specify an alternative location using the `tempdir` option.
 */
struct workbook_options {
    workbook_options() : constant_memory(false) {}

    /** Optimize the workbook to use constant memory for worksheets */
    bool constant_memory;

    /** Directory to use for the temporary files created by libxlsxwriter. */
    std::string tmpdir;
};

class packager;

/**
 * @brief Struct to represent an Excel workbook.
 *
 * The members of the lxw_workbook struct aren't modified directly. Instead
 * the workbook properties are set by calling the functions shown in
 * workbook.h.
 */
class workbook : public xmlwriter {
    friend class packager;

public:
    /**
     * @brief Create a new workbook object, and set the workbook options.
     *
     * @param filename The name of the new Excel file to create.
     * @param options  Workbook options.
     *
     *
     * @code
     *    workbook_options options;
     *    options.constant_memory = 1;
     *    options.tmpdir = "C:\\Temp";
     *
     *    workbook_ptr workbook  = std::make_shared<workbook>("filename.xlsx", options);
     * @endcode
     *
     * The options that can be set via #lxw_workbook_options are:
     *
     * - `constant_memory`: Reduces the amount of data stored in memory so that
     *   large files can be written efficiently.
     *
     *   @note In this mode a row of data is written and then discarded when a
     *   cell in a new row is added via one of the `worksheet_write_*()`
     *   methods. Therefore, once this option is active, data should be written in
     *   sequential row order. For this reason the `worksheet_merge_range()`
     *   doesn't work in this mode. See also @ref ww_mem_constant.
     *
     * - `tmpdir`: libxlsxwriter stores workbook data in temporary files prior
     *   to assembling the final XLSX file. The temporary files are created in the
     *   system's temp directory. If the default temporary directory isn't
     *   accessible to your application, or doesn't contain enough space, you can
     *   specify an alternative location using the `tempdir` option.*
     *
     * See @ref working_with_memory for more details.
     *
     */

    workbook(const std::string& filename, const workbook_options &options = workbook_options());

    ~workbook();

    /**
     * @brief get_worksheets
     * @return list of worksheets
     */
    std::vector<worksheet_ptr> get_worksheets();

    /**
     * @brief Add a new worksheet to a workbook:
     *
     * @param sheetname Optional worksheet name, defaults to Sheet1, etc.
     *
     * @return A worksheet_ptr object.
     *
     * The `%add_worksheet()` function adds a new worksheet to a workbook:
     *
     * At least one worksheet should be added to a new workbook: The @ref
     * worksheet.h "Worksheet" object is used to write data and configure a
     * worksheet in the workbook.
     *
     * The `sheetname` parameter is optional. If it is `NULL` the default
     * Excel convention will be followed, i.e. Sheet1, Sheet2, etc.:
     *
     * @code
     *     worksheet = workbook->add_worksheet();           // Sheet1
     *     worksheet = workbook->add_worksheet("Foglio2");  // Foglio2
     *     worksheet = workbook->add_worksheet("Data");     // Data
     *     worksheet = workbook->add_worksheet();           // Sheet4
     *
     * @endcode
     *
     * @image html workbook02.png
     *
     * The worksheet name must be a valid Excel worksheet name, i.e. it must be
     * less than 32 character and it cannot contain any of the characters:
     *
     *     / \ [ ] : * ?
     *
     * In addition, you cannot use the same, case insensitive, `sheetname` for more
     * than one worksheet.
     *
     */
     worksheet_ptr add_worksheet(const std::string& sheetname = std::string());

    /**
     * @brief Create a new @ref format.hpp "Format" object to formats cells in
     *        worksheets.
     *
     * @return A lxw_format instance.
     *
     * The `workbook->add_format()` function can be used to create new @ref
     * format.hpp "Format" objects which are used to apply formatting to a cell.
     *
     * @code
     *    // Create the Format.
     *    format_ptr format = workbook->add_format();
     *
     *    // Set some of the format properties.
     *    format->set_bold();
     *    format->set_font_color(LXW_COLOR_RED);
     *
     *    // Use the format to change the text format in a cell.
     *    worksheet->write_string(0, 0, "Hello", format);
     * @endcode
     *
     * See @ref format.h "the Format object" and @ref working_with_formats
     * sections for more details about Format properties and how to set them.
     *
     */
    format_ptr add_format();

    /**
     * @brief Create a new chart to be added to a worksheet:
     *
     * @param chart_type The type of chart to be created. See #chart_types.
     *
     * @return A chart object.
     *
     * The `%add_chart()` function creates a new chart object that can
     * be added to a worksheet:
     *
     * @code
     *     // Create a chart object.
     *     chart_ptr chart = workbook->add_chart(LXW_CHART_COLUMN);
     *
     *     // Add data series to the chart.
     *     chart->add_series(NULL, "Sheet1!$A$1:$A$5");
     *     chart->add_series(NULL, "Sheet1!$B$1:$B$5");
     *     chart->add_series(NULL, "Sheet1!$C$1:$C$5");
     *
     *     // Insert the chart into the worksheet
     *     worksheet->insert_chart(CELL("B7"), chart);
     * @endcode
     *
     * The available chart types are defined in #chart_types. The types of
     * charts that are supported are:
     *
     * | Chart type                               | Description                            |
     * | :--------------------------------------- | :------------------------------------  |
     * | #LXW_CHART_AREA                          | Area chart.                            |
     * | #LXW_CHART_AREA_STACKED                  | Area chart - stacked.                  |
     * | #LXW_CHART_AREA_STACKED_PERCENT          | Area chart - percentage stacked.       |
     * | #LXW_CHART_BAR                           | Bar chart.                             |
     * | #LXW_CHART_BAR_STACKED                   | Bar chart - stacked.                   |
     * | #LXW_CHART_BAR_STACKED_PERCENT           | Bar chart - percentage stacked.        |
     * | #LXW_CHART_COLUMN                        | Column chart.                          |
     * | #LXW_CHART_COLUMN_STACKED                | Column chart - stacked.                |
     * | #LXW_CHART_COLUMN_STACKED_PERCENT        | Column chart - percentage stacked.     |
     * | #LXW_CHART_DOUGHNUT                      | Doughnut chart.                        |
     * | #LXW_CHART_LINE                          | Line chart.                            |
     * | #LXW_CHART_PIE                           | Pie chart.                             |
     * | #LXW_CHART_SCATTER                       | Scatter chart.                         |
     * | #LXW_CHART_SCATTER_STRAIGHT              | Scatter chart - straight.              |
     * | #LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS | Scatter chart - straight with markers. |
     * | #LXW_CHART_SCATTER_SMOOTH                | Scatter chart - smooth.                |
     * | #LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS   | Scatter chart - smooth with markers.   |
     * | #LXW_CHART_RADAR                         | Radar chart.                           |
     * | #LXW_CHART_RADAR_WITH_MARKERS            | Radar chart - with markers.            |
     * | #LXW_CHART_RADAR_FILLED                  | Radar chart - filled.                  |
     *
     *
     *
     * See @ref chart.h for details.
     */
    chart_ptr add_chart(uint8_t chart_type);

    /**
     * @brief Close the Workbook object and write the XLSX file.
     *
     * @return A #lxw_error.
     *
     * The `%close()` function closes a Workbook object, writes the Excel
     * file to disk, frees any memory allocated internally to the Workbook and
     * frees the object itself.
     *
     * @code
     *     workbook->close();
     * @endcode
     *
     * The `%close()` function returns any #lxw_error error codes
     * encountered when creating the Excel file. The error code can be returned
     * from the program main or the calling function:
     *
     * @code
     *     return workbook->close();
     * @endcode
     *
     */
    lxw_error close();

    /**
     * @brief Set the document properties such as Title, Author etc.
     *
     * @param properties Document properties to set.
     *
     * @return A #lxw_error.
     *
     * The `%set_properties` method can be used to set the document
     * properties of the Excel file created by `libxlsxwriter`. These properties
     * are visible when you use the `Office Button -> Prepare -> Properties`
     * option in Excel and are also available to external applications that read
     * or index windows files.
     *
     * The properties that can be set are:
     *
     * - `title`
     * - `subject`
     * - `author`
     * - `manager`
     * - `company`
     * - `category`
     * - `keywords`
     * - `comments`
     * - `hyperlink_base`
     *
     * The properties are specified via a `doc_properties` struct. All the
     * members are `std::string` and they are all optional. An example of how to create
     * and pass the properties is:
     *
     * @code
     *     // Create a properties structure and set some of the fields.
     *     doc_properties properties;
     *     properties.title    = "This is an example spreadsheet";
     *     properties.subject  = "With document properties";
     *     properties.author   = "John McNamara";
     *     properties.manager  = "Dr. Heinz Doofenshmirtz";
     *     properties.company  = "of Wolves";
     *     properties.category = "Example spreadsheets";
     *     properties.keywords = "Sample, Example, Properties";
     *     properties.comments = "Created with libxlsxwriter";
     *     properties.status   = "Quo";
     *
     *     // Set the properties in the workbook.
     *     workbook->set_properties(properties);
     * @endcode
     *
     * @image html doc_properties.png
     *
     */
    lxw_error set_properties(const doc_properties& properties);

    /**
     * @brief Set a custom document text property.
     *
     * @param name     The name of the custom property.
     * @param value    The value of the custom property.
     *
     * @return A #lxw_error.
     *
     * The `%workbook_set_custom_property_string()` method can be used to set one
     * or more custom document text properties not covered by the standard
     * properties in the `workbook_set_properties()` method above.
     *
     *  For example:
     *
     * @code
     *     workbook->set_custom_property_string("Checked by", "Eve");
     * @endcode
     *
     * @image html custom_properties.png
     *
     * There are 4 `workbook_set_custom_property_string_*()` functions for each
     * of the custom property types supported by Excel:
     *
     * - text/string: `set_custom_property_string()`
     * - number:      `set_custom_property_number()`
     * - datetime:    `set_custom_property_datetime()`
     * - boolean:     `set_custom_property_boolean()`
     *
     * **Note**: the name and value parameters are limited to 255 characters
     * by Excel.
     *
     */
    lxw_error set_custom_property_string(const std::string& name,
                                         const std::string& value);
    /**
     * @brief Set a custom document number property.
     *
     * @param name     The name of the custom property.
     * @param value    The value of the custom property.
     *
     * @return A #lxw_error.
     *
     * Set a custom document number property.
     * See `set_custom_property_string()` above for details.
     *
     * @code
     *     workbook->set_custom_property_number("Document number", 12345);
     * @endcode
     */
    lxw_error set_custom_property_number(const std::string& name, double value);

    /* Undocumented since the user can use set_custom_property_number().
     * Only implemented for file format completeness and testing.
     */
    lxw_error set_custom_property_integer(const std::string& name, int32_t value);

    /**
     * @brief Set a custom document boolean property.
     *
     * @param name     The name of the custom property.
     * @param value    The value of the custom property.
     *
     * @return A #lxw_error.
     *
     * Set a custom document boolean property.
     * See `set_custom_property_string()` above for details.
     *
     * @code
     *     workbook->set_custom_property_boolean("Has Review", 1);
     * @endcode
     */
    lxw_error set_custom_property_boolean(const std::string& name, bool value);
    /**
     * @brief Set a custom document date or time property.
     *
     * @param name     The name of the custom property.
     * @param datetime The value of the custom property.
     *
     * @return A #lxw_error.
     *
     * Set a custom date or time number property.
     * See `workbook_set_custom_property_string()` above for details.
     *
     * @code
     * @todo
     *     lxw_datetime datetime  = {2016, 12, 1,  11, 55, 0.0};
     *
     *     workbook->set_custom_property_datetime("Date completed", &datetime);
     * @endcode
     */
    lxw_error set_custom_property_datetime(const std::string& name, lxw_datetime *datetime);

    /**
     * @brief Create a defined name in the workbook to use as a variable.
     *
     * @param name     The defined name.
     * @param formula  The cell or range that the defined name refers to.
     *
     * @return A #lxw_error.
     *
     * This method is used to defined a name that can be used to represent a
     * value, a single cell or a range of cells in a workbook: These defined names
     * can then be used in formulas:
     *
     * @code
     *     workbook->define_name("Exchange_rate", "=0.96");
     *     worksheet->write_formula(2, 1, "=Exchange_rate", NULL);
     *
     * @endcode
     *
     * @image html defined_name.png
     *
     * As in Excel a name defined like this is "global" to the workbook and can be
     * referred to from any worksheet:
     *
     * @code
     *     // Global workbook name.
     *     workbook->define_name("Sales", "=Sheet1!$G$1:$H$10");
     * @endcode
     *
     * It is also possible to define a local/worksheet name by prefixing it with
     * the sheet name using the syntax `'sheetname!definedname'`:
     *
     * @code
     *     // Local worksheet name.
     *     workbook->define_name("Sheet2!Sales", "=Sheet2!$G$1:$G$10");
     * @endcode
     *
     * If the sheet name contains spaces or special characters you must follow the
     * Excel convention and enclose it in single quotes:
     *
     * @code
     *     workbook->define_name("'New Data'!Sales", "=Sheet2!$G$1:$G$10");
     * @endcode
     *
     * The rules for names in Excel are explained in the
     * [Microsoft Office
    documentation](http://office.microsoft.com/en-001/excel-help/define-and-use-names-in-formulas-HA010147120.aspx).
     *
     */
    lxw_error define_name(const std::string& name, const std::string& formula);

    worksheet_ptr get_worksheet_by_name(const std::string& name);

    void assemble_xml_file();
    void set_default_xf_indices();

private:
    FILE *file;
    std::vector<worksheet_ptr> worksheets;
    std::map<std::string, worksheet_ptr> worksheet_names;
    std::vector<chart_ptr> charts;
    std::vector<chart_ptr> ordered_charts;
    std::vector<format_ptr> formats;
    std::vector<defined_name_ptr> defined_names;
    sst_ptr sst;
    doc_properties properties;
    std::list<custom_property_ptr> custom_properties;

    std::string filename;
    workbook_options options;

    uint16_t num_sheets;
    uint16_t first_sheet;
    uint16_t active_sheet;
    uint16_t num_xf_formats;
    uint16_t num_format_count;
    uint16_t drawing_count;

    uint16_t font_count;
    uint16_t border_count;
    uint16_t fill_count;
    uint8_t optimize;

    bool has_png;
    bool has_jpeg;
    bool has_bmp;

    //lxw_hash_table *used_xf_formats;
    std::unordered_set<std::shared_ptr<format>> used_xf_formats;

    /* Declarations required for unit testing. */
#ifdef TESTING

    void _xml_declaration();
    void _write_workbook();
    void _write_file_version();
    void _write_workbook_pr();
    void _write_book_views();
    void _write_workbook_view();
    void _write_sheet(const std::string& name, uint32_t sheet_id, bool hidden);
    void _write_sheets();
    void _write_calc_pr();

    void _write_defined_name(lxw_defined_name *define_name);
    void _write_defined_names();

    lxw_error _store_defined_name(const std::string& name,
                                  const std::string& app_name,
                                  const std::string& formula, int16_t index,
                                  bool hidden);

#endif /* TESTING */

};

typedef std::shared_ptr<workbook> workbook_ptr;

} // namespace xlsxwriter

#endif /* __LXW_WORKBOOK_H__ */

