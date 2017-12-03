/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * chart - A libxlsxwriter library for creating Excel XLSX chart files.
 *
 */

/**
 * @page chart_page The Chart object
 *
 * The Chart object represents an Excel chart. It provides functions for
 * adding data series to the chart and for configuring the chart.
 *
 * See @ref chart.h for full details of the functionality.
 *
 * @file chart.h
 *
 * @brief Functions related to adding data to and configuring  a chart.
 *
 * The Chart object represents an Excel chart. It provides functions for
 * adding data series to the chart and for configuring the chart.
 *
 * A Chart object isn't created directly. Instead a chart is created by
 * calling the `workbook_add_chart()` function from a Workbook object. For
 * example:
 *
 * @code
 *
 * #include "xlsxwriter.hpp"
 *
 * int main() {
 *
 *     std::shared_ptr<xlsxwriter::workbook> workbook = std::make_shared<xlsxwriter::workbook>("chart.xlsx");
 *     xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
 *
 *     // User function to add data to worksheet, not shown here.
 *     write_worksheet_data(worksheet);
 *
 *     // Create a chart object.
 *     xlsxwriter::chart *chart = workbook->add_chart( LXW_CHART_COLUMN);
 *
 *     // In the simplest case we just add some value data series.
 *     // The NULL categories will default to 1 to 5 like in Excel.
 *     chart->add_series("", "=Sheet1!$A$1:$A$5");
 *     chart->add_series("", "=Sheet1!$B$1:$B$5");
 *     chart->add_series("", "=Sheet1!$C$1:$C$5");
 *
 *     // Insert the chart into the worksheet
 *     worksheet->insert_chart(CELL("B7"), chart);
 *
 *     int result = workbook->close(); return result;
 * }
 *
 * @endcode
 *
 * The chart in the worksheet will look like this:
 * @image html chart_simple.png
 *
 * The basic procedure for adding a chart to a worksheet is:
 *
 * 1. Create the chart with `workbook_add_chart()`.
 * 2. Add one or more data series to the chart which refers to data in the
 *    workbook using `chart_add_series()`.
 * 3. Configure the chart with the other available functions shown below.
 * 4. Insert the chart into a worksheet using `worksheet_insert_chart()`.
 *
 */

#ifndef __LXW_CHART_H__
#define __LXW_CHART_H__

#include <stdint.h>
#include <string.h>

#include <memory>
#include <vector>
#include "xmlwriter.hpp"
#include "common.hpp"
#include "shape.hpp"

namespace xlsxwriter {

struct val_axis_args;

#define LXW_CHART_NUM_FORMAT_LEN 128

/** Available chart types . */
typedef enum chart_types {

    /** None. */
    LXW_CHART_NONE = 0,

    /** Area chart. */
    LXW_CHART_AREA,

    /** Area chart - stacked. */
    LXW_CHART_AREA_STACKED,

    /** Area chart - percentage stacked. */
    LXW_CHART_AREA_STACKED_PERCENT,

    /** Bar chart. */
    LXW_CHART_BAR,

    /** Bar chart - stacked. */
    LXW_CHART_BAR_STACKED,

    /** Bar chart - percentage stacked. */
    LXW_CHART_BAR_STACKED_PERCENT,

    /** Column chart. */
    LXW_CHART_COLUMN,

    /** Column chart - stacked. */
    LXW_CHART_COLUMN_STACKED,

    /** Column chart - percentage stacked. */
    LXW_CHART_COLUMN_STACKED_PERCENT,

    /** Doughnut chart. */
    LXW_CHART_DOUGHNUT,

    /** Line chart. */
    LXW_CHART_LINE,

    /** Pie chart. */
    LXW_CHART_PIE,

    /** Scatter chart. */
    LXW_CHART_SCATTER,

    /** Scatter chart - straight. */
    LXW_CHART_SCATTER_STRAIGHT,

    /** Scatter chart - straight with markers. */
    LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS,

    /** Scatter chart - smooth. */
    LXW_CHART_SCATTER_SMOOTH,

    /** Scatter chart - smooth with markers. */
    LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS,

    /** Radar chart. */
    LXW_CHART_RADAR,

    /** Radar chart - with markers. */
    LXW_CHART_RADAR_WITH_MARKERS,

    /** Radar chart - filled. */
    LXW_CHART_RADAR_FILLED
} chart_types;

enum chart_subtypes {

    LXW_CHART_SUBTYPE_NONE = 0,
    LXW_CHART_SUBTYPE_STACKED,
    LXW_CHART_SUBTYPE_STACKED_PERCENT
};

enum chart_groupings {
    LXW_GROUPING_CLUSTERED,
    LXW_GROUPING_STANDARD,
    LXW_GROUPING_PERCENTSTACKED,
    LXW_GROUPING_STACKED
};

enum chart_axis_positions {
    LXW_CHART_AXIS_POSITION_BETWEEN,
    LXW_CHART_AXIS_POSITION_ON_TICK
};

enum chart_positions {
    LXW_CHART_RIGHT,
    LXW_CHART_LEFT,
    LXW_CHART_TOP,
    LXW_CHART_BOTTOM
};

enum lxw_marker_types {
    LXW_MARKER_NONE = 0,
    LXW_MARKER_AUTOMATIC,
    LXW_MARKER_SQUARE,
    LXW_MARKER_DIAMOND,
    LXW_MARKER_TRIANGLE,
    LXW_MARKER_X,
    LXW_MARKER_STAR,
    LXW_MARKER_DOT,
    LXW_MARKER_SHORT_DASH,
    LXW_MARKER_DASH,
    LXW_MARKER_LONG_DASH,
    LXW_MARKER_CIRCLE,
    LXW_MARKER_PLUS,
    LXW_MARKER_PICTURE
};

struct series_data_point {
    bool is_string;
    double number;
    std::string* string;
    bool no_data;
};

struct series_range {
    series_range();
    std::string formula;
    std::string sheetname;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
    bool ignore_cache;

    bool has_string_cache;
    uint16_t num_data_points;
    std::vector<std::shared_ptr<series_data_point>> data_cache;
};

typedef std::shared_ptr<series_range> series_range_ptr;

struct chart_font {

    uint8_t bold;

};

struct XLSXWRITER_EXPORT chart_title {

    chart_title();

    std::string name;
    lxw_row_t row;
    lxw_col_t col;
    chart_font font;
    int32_t angle;
    bool off;
    bool is_horizontal;
    bool ignore_cache;

    /* We use a range to hold the title formula properties even though it
     * will only have 1 point in order to re-use similar functions.*/
    series_range_ptr range;

    series_data_point data_point;

};

struct lxw_marker {
    uint8_t marker_type;
    lxw_shape_properties properties;
};

class workbook;
class worksheet;

/**
 * @brief Struct to represent an Excel chart data series.
 *
 * The chart_series is created using the chart_add_series function. It is
 * used in functions that modify a chart series but the members of the struct
 * aren't modified directly.
 */
struct XLSXWRITER_EXPORT chart_series {
    friend class chart;
    friend class workbook;
    friend class worksheet;
public:

    chart_series();

    /**
     * @brief Set a series "categories" range using row and column values.
     *
     * @param sheetname The name of the worksheet that contains the data range.
     * @param first_row The first row of the range. (All zero indexed.)
     * @param first_col The first column of the range.
     * @param last_row  The last row of the range.
     * @param last_col  The last col of the range.
     *
     * The `categories` and `values` of a chart data series are generally set
     * using the `chart_add_series()` function and Excel range formulas like
     * `"=Sheet1!$A$2:$A$7"`.
     *
     * The `%chart_series_set_categories()` function is an alternative method that
     * is easier to generate programmatically. It requires that you set the
     * `categories` and `values` parameters in `chart_add_series()`to `NULL` and
     * then set them using row and column values in
     * `chart_series_set_categories()` and `series->set_values()`:
     *
     * @code
     *     chart_series_ptr series = chart->add_series();
     *
     *     // Configure the series ranges programmatically.
     *     series->set_categories("Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
     *     series->set_values("Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
     * @endcode
     *
     */
    void set_categories(const std::string& sheetname, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);

    /**
     * @brief Set a series "values" range using row and column values.
     *
     * @param series    A series object created via `chart_add_series()`.
     * @param sheetname The name of the worksheet that contains the data range.
     * @param first_row The first row of the range. (All zero indexed.)
     * @param first_col The first column of the range.
     * @param last_row  The last row of the range.
     * @param last_col  The last col of the range.
     *
     * The `categories` and `values` of a chart data series are generally set
     * using the `chart_add_series()` function and Excel range formulas like
     * `"=Sheet1!$A$2:$A$7"`.
     *
     * The `%series->set_values()` function is an alternative method that is
     * easier to generate programmatically. See the documentation for
     * `chart_series_set_categories()` above.
     */
    void set_values(const std::string& sheetname, lxw_row_t first_row, lxw_col_t first_col,
                    lxw_row_t last_row, lxw_col_t last_col);

    /**
     * @brief Set the name of a chart series range.
     *
     * @param name   The series name.
     *
     * The `%set_name` function is used to set the name for a chart
     * data series. The series name in Excel is displayed in the chart legend and
     * in the formula bar. The name property is optional and if it isn't supplied
     * it will default to `Series 1..n`.
     *
     * The function applies to a #chart_series object created using
     * `add_series()`:
     *
     * @code
     *     chart_series_ptr series = chart->add_series("", "=Sheet1!$B$2:$B$7");
     *
     *     series->set_name("Quarterly budget data");
     * @endcode
     *
     * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
     * a cell in the workbook that contains the name:
     *
     * @code
     *     chart_series_ptr series = chart->add_series("", "=Sheet1!$B$2:$B$7");
     *
     *     series->set_name("=Sheet1!$B1$1");
     * @endcode
     *
     * See also the `set_name_range()` function to see how to set the
     * name formula programmatically.
     */
    void set_name(const std::string& name);

    /**
     * @brief Set a series name formula using row and column values.
     *
     * @param sheetname The name of the worksheet that contains the cell range.
     * @param row       The zero indexed row number of the range.
     * @param col       The zero indexed column number of the range.
     *
     * The `%set_name_range()` function can be used to set a series
     * name range and is an alternative to using `set_name()` and a
     * string formula:
     *
     * @code
     *     chart_series_ptr series = chart->add_series("", "=Sheet1!$B$2:$B$7");
     *
     *     series->set_name_range("Sheet1", 0, 2); // "=Sheet1!$C$1"
     * @endcode
     */
    void set_name_range(const std::string& sheetname, lxw_row_t row, lxw_col_t col);

    /// TODO make this private
//private:

    series_range_ptr categories;
    series_range_ptr values;
    chart_title title;
    lxw_shape_properties properties;
    lxw_marker marker;
    bool x2_axis;
    bool y2_axis;
};

typedef std::shared_ptr<chart_series> chart_series_ptr;

/**
 * @brief Struct to represent an Excel chart axis. It is used in functions
 * that modify a chart axis but the members of the struct aren't modified
 * directly.
 */
struct XLSXWRITER_EXPORT chart_axis{
    friend class chart;
    friend class workbook;
public:

    chart_axis();

    /**
     * @brief Set the name caption of the an axis.
     *
     * @param axis A pointer to a chart #chart_axis object.
     * @param name The name caption of the axis.
     *
     * The `%chart_axis_set_name()` function sets the name (also known as title or
     * caption) for an axis. It can be used for the X or Y axes. The name is
     * displayed below an X axis and to the side of a Y axis.
     *
     * @code
     *     chart->x1_axis()->set_name("Earnings per Quarter");
     *     chart->y1_axis()->axis_set_name("US Dollars (Millions)");
     * @endcode
     *
     * @image html chart_axis_set_name.png
     *
     * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
     * a cell in the workbook that contains the name:
     *
     * @code
     *     chart->x1_axis()->set_name("=Sheet1!$B1$1");
     * @endcode
     *
     * See also the `chart_axis_set_name_range()` function to see how to set the
     * name formula programmatically.
     *
     * This function is applicable to category, date and value axes.
     */
    void set_name(const std::string& name);

    /**
     * @brief Set a chart axis name formula using row and column values.
     *
     * @param axis      A pointer to a chart #chart_axis object.
     * @param sheetname The name of the worksheet that contains the cell range.
     * @param row       The zero indexed row number of the range.
     * @param col       The zero indexed column number of the range.
     *
     * The `%chart_axis_set_name_range()` function can be used to set an axis name
     * range and is an alternative to using `chart_axis_set_name()` and a string
     * formula:
     *
     * @code
     *     chart->x1_axis()->set_name_range("Sheet1", 1, 0);
     *     chart->y1_axis()->set_name_range("Sheet1", 2, 0);
     * @endcode
     */
    void set_name_range(const std::string& sheetname, lxw_row_t row, lxw_col_t col);

    /**
     * @brief Set a chart axis values format
     *
     * @param axis     A pointer to a chart #chart_axis object.
     * @param format   Format for axis's values
     */
    void set_format(const std::string& format);

    void set_crossing(const std::string& crossing);

    void set_major_tick_mark(bool mark);

    void set_default_num_format(const std::string &format);

    void set_default_major_gridlines(bool mark);
private:

    chart_title title;

    std::string num_format;
    std::string default_num_format;
    std::string crossing;

    bool default_major_gridlines;
    bool major_tick_mark;

    double min_value;
    double max_value;

    uint8_t position;
    bool visible;

};

struct series_options {
    uint8_t x2_axis;
    uint8_t y2_axis;
};

class packager;
class workbook;
class worksheet;

/**
 * @brief Class to represent an Excel chart.
 *
 * The members of the chart class aren't modified directly. Instead
 * the chart properties are set by calling the functions shown in chart.h.
 */
class XLSXWRITER_EXPORT chart : public xmlwriter {
    friend class packager;
    friend class workbook;
    friend class worksheet;
public:

    chart(uint8_t type);

    virtual ~chart();

    void assemble_xml_file();


    /**
     * @brief Add a data series to a chart.
     *
     * @param categories The range of categories in the data series.
     * @param values     The range of values in the data series.
     *
     * @return A chart_series object pointer.
     *
     * In Excel a chart **series** is a collection of information that defines
     * which data is plotted such as the categories and values. It is also used to
     * define the formatting for the data.
     *
     * For an libxlsxwriter chart object the `%add_series()` function is
     * used to set the categories and values of the series:
     *
     * @code
     *     chart->add_series("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
     * @endcode
     *
     *
     * The series parameters are:
     *
     * - `categories`: This sets the chart category labels. The category is more
     *   or less the same as the X axis. In most Excel chart types the
     *   `categories` property is optional and the chart will just assume a
     *   sequential series from `1..n`:
     *
     * @code
     *     // The NULL category will default to 1 to 5 like in Excel.
     *     add_series("", "Sheet1!$A$1:$A$5");
     * @endcode
     *
     *  - `values`: This is the most important property of a series and is the
     *    only mandatory option for every chart object. This parameter links the
     *    chart with the worksheet data that it displays.
     *
     * The `categories` and `values` should be a string formula like
     * `"=Sheet1!$A$2:$A$7"` in the same way it is represented in Excel. This is
     * convenient when recreating a chart from an example in Excel but it is
     * trickier to generate programmatically. For these cases you can set the
     * `categories` and `values` to `NULL` and use the
     * `chart_series_set_categories()` and `series->set_values()` functions:
     *
     * @code
     *     chart_series_ptr series = chart->add_series(chart);
     *
     *     // Configure the series using a syntax that is easier to define programmatically.
     *     chart_series_ptr->set_categories("Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
     *     chart_series_ptr->set_values("Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
     * @endcode
     *
     * As shown in the previous example the return value from
     * `%add_series()` is a chart_series pointer. This can be used in
     * other functions that configure a series.
     *
     *
     * More than one series can be added to a chart. The series numbering and
     * order in the Excel chart will be the same as the order in which they are
     * added in libxlsxwriter:
     *
     * @code
     *    chart->add_series("", "Sheet1!$A$1:$A$5");
     *    chart->add_series("", "Sheet1!$B$1:$B$5");
     *    chart->add_series("", "Sheet1!$C$1:$C$5");
     * @endcode
     *
     * It is also possible to specify non-contiguous ranges:
     *
     * @code
     *    chart->add_series(
     *        "=(Sheet1!$A$1:$A$9,Sheet1!$A$14:$A$25)",
     *        "=(Sheet1!$B$1:$B$9,Sheet1!$B$14:$B$25)"
     *    );
     * @endcode
     *
     */
    virtual chart_series* add_series(const std::string& categories = std::string(), const std::string& values = std::string(), const series_options& options = series_options());

    void chart_set_y2_axis(const std::shared_ptr<chart_axis>& axis);

    /**
     * @brief Set the title of the chart.
     *
     * @param name  The chart title name.
     *
     * The `%title_set_name()` function sets the name (title) for the
     * chart. The name is displayed above the chart.
     *
     * @code
     *     chart->title_set_name("Year End Results");
     * @endcode
     *
     * @image html chart_title_set_name.png
     *
     * The name parameter can also be a formula such as `=Sheet1!$A$1` to point to
     * a cell in the workbook that contains the name:
     *
     * @code
     *     chart->title_set_name("=Sheet1!$B1$1");
     * @endcode
     *
     * See also the `title_set_name_range()` function to see how to set the
     * name formula programmatically.
     *
     * The Excel default is to have no chart title.
     */
    void title_set_name(const std::string& name);

    /**
     * @brief Set a chart title formula using row and column values.
     *
     * @param sheetname The name of the worksheet that contains the cell range.
     * @param row       The zero indexed row number of the range.
     * @param col       The zero indexed column number of the range.
     *
     * The `%title_set_name_range()` function can be used to set a chart
     * title range and is an alternative to using `title_set_name()` and a
     * string formula:
     *
     * @code
     *     chart->title_set_name_range("Sheet1", 1, 0);
     * @endcode
     */
    void title_set_name_range(const std::string& sheetname, lxw_row_t row, lxw_col_t col);

    /**
     * @brief Turn off an automatic chart title.
     *
     * In general in Excel a chart title isn't displayed unless the user
     * explicitly adds one. However, Excel adds an automatic chart title to charts
     * with a single series and a user defined series name. The
     * `chart_title_off()` function allows you to turn off this automatic chart
     * title:
     *
     * @code
     *     chart->title_off(chart);
     * @endcode
     */
    void title_off();

    chart_axis* get_x_axis();
    chart_axis* get_y_axis();

    /**
     * @brief Set the chart style type.
     *
     * @param chart    Pointer to a xlsxwriter::chart instance to be configured.
     * @param style_id An index representing the chart style, 1 - 48.
     *
     * The `%chart_set_style()` function is used to set the style of the chart to
     * one of the 48 built-in styles available on the "Design" tab in Excel 2007:
     *
     * @code
     *     chart->set_style(37)
     * @endcode
     *
     * @image html chart_style.png
     *
     * The style index number is counted from 1 on the top left in the Excel
     * dialog. The default style is 2.
     *
     * **Note:**
     *
     * In Excel 2013 the Styles section of the "Design" tab in Excel shows what
     * were referred to as "Layouts" in previous versions of Excel. These layouts
     * are not defined in the file format. They are a collection of modifications
     * to the base chart type. They can not be defined by the `chart_set_style()``
     * function.
     *
     *
     */
    void set_style(uint8_t style_id);

    void set_rotation(uint16_t rotation);
    void set_hole_size(uint8_t size);

    static void set_range(const series_range_ptr& range, const std::string &sheetname, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);

    void set_y2_axis(const std::shared_ptr<chart_axis> &axis);

protected:

    uint8_t type;
    uint8_t subtype;
    uint16_t series_index;

    virtual void write_chart_type(bool) = 0;
    virtual void write_plot_area();
    virtual void _initialize() = 0;

    /**
     * A pointer to the chart x_axis object which can be used in functions
     * that configures the X axis.
     */
    std::shared_ptr<chart_axis> x_axis;

    /**
     * A pointer to the chart y_axis object which can be used in functions
     * that configures the Y axis.
     */
    std::shared_ptr<chart_axis> y_axis;

    std::shared_ptr<chart_axis> x2_axis;

    std::shared_ptr<chart_axis> y2_axis;

    chart_title title;

    uint32_t id;

    /// this stupid piece of crap is written only because of functional tests
    /// TODO: get rid of this
public:
    uint32_t axis_id_1;
    uint32_t axis_id_2;

protected:

    uint32_t axis_id_3;
    uint32_t axis_id_4;

    bool in_use;
    bool is_scatter;
    bool cat_has_num_fmt;

    uint8_t has_horiz_cat_axis;
    uint8_t has_horiz_val_axis;

    uint8_t style_id;
    uint16_t rotation;
    uint16_t hole_size;

    bool no_title;
    bool has_markers;
    bool has_overlap;
    int series_overlap_1;

    uint8_t grouping;
    uint8_t cross_between;
    uint8_t cat_axis_position;
    uint8_t val_axis_position;
	uint8_t legend_position;
    bool is_secondary;
    std::shared_ptr<chart> combined;

    std::vector<chart_series_ptr> series_list;

    void _write_chart_space();
    void _write_lang();
    void _write_layout();
    void _write_vary_colors();
    void _write_radar_style();
    void _write_grouping(uint8_t grouping);
    void _write_first_slice_ang();
    void _write_hole_size();
    void _write_a_t(const std::string &name);
    void _write_a_end_para_rpr();
    void _write_a_def_rpr();
    void _write_a_r_pr();
    void _write_a_r(const std::string &name);
    void _write_a_p_pr();
    void _write_a_p_pr_pie();
    void _add_axis_ids(bool primary);
    void _write_a_p_pr_rich();
    void _write_a_p();
    void _write_a_p_pie();
    void _write_a_p_rich(const std::string &name);
    void _write_a_lst_style();
    void _write_a_body_pr(chart_title *title);
    void _write_pt_count(uint16_t num_data_points);
    void _write_v_num(double number);
    void _write_v_str(const std::string &str);
    void _write_pt(uint16_t index, const std::shared_ptr<series_data_point>& data_point);
    void _write_f(const std::string &formula);
    void _write_num_pt(uint16_t index, const std::shared_ptr<series_data_point>& data_point);
    void _write_format_code();
    void _write_num_cache(const series_range_ptr&range);
    void _write_str_cache(const series_range_ptr& range);
    void _write_num_ref(const series_range_ptr& range);
    void _write_data_cache(const series_range_ptr& range, bool has_string_cache);
    void _write_str_ref(const series_range_ptr& range);
    void _write_tx_value(const std::string &name);
    void _write_tx_formula(chart_title *title);
    void _write_tx_pr(chart_title *title);
    void _write_rich(chart_title *title);
    void _write_tx_pr_pie(chart_title *title);
    void _write_tx_rich(chart_title *title);
    void _write_title_rich(chart_title *title);
    void _write_title_formula(chart_title *title);
    void _write_auto_title_deleted();
    void _write_idx(uint16_t index);
    void _write_a_alpha(double transparency);
    void _write_a_no_fill();
    void _write_a_solid_fill(lxw_color_t color, double transparency);
    void _write_a_ln(lxw_line *line);
    void _write_sp_pr(lxw_shape_properties *properties);
    void _write_order(uint16_t index);
    void _write_axis_id(uint32_t axis_id);
    void _write_axis_ids(uint8_t primary);
    void _write_series_name(const std::shared_ptr<chart_series> &series);
    void _write_major_tick_mark(const std::shared_ptr<chart_axis> &axis);
    void _write_symbol();
    void _write_marker_data(lxw_marker *marker);
    void _write_marker(lxw_marker *marker);
    void _write_marker_value();
    void _write_smooth();
    void _write_scatter_style();
    void _write_cat(const std::shared_ptr<chart_series> &series);
    void _write_x_val(const std::shared_ptr<chart_series> &series);
    void _write_y_val(const std::shared_ptr<chart_series> &series);
    void _write_val(const std::shared_ptr<chart_series> &series);
    void _write_style();
    void _write_a_srgb(lxw_color_t color, double transparency);
    void _write_radar_chart(uint8_t primary_axes);
    void _write_orientation();
    void _write_tick_lbl_pos();
    void _write_axis_pos(uint8_t position);
    void _write_ser(const std::shared_ptr<chart_series> &series);
    void _write_lbl_offset();
    void _write_lbl_algn();
    void _write_scaling(const std::shared_ptr<chart_axis> &axis);
    void _write_cross_axis(uint32_t axis_id);
    void _write_crosses(const std::string &value);
    void _write_auto();
    void _write_major_gridlines(const std::shared_ptr<chart_axis> &axis);
    void _write_number_format(const std::shared_ptr<chart_axis> &axis);
    void _write_cross_between();
    void _write_legend_pos();
    void _write_legend();
    void _write_plot_vis_only();
    void _write_header_footer();
    void _write_page_margins();
    void _write_page_setup();
    void _write_overlap(int overlap);
    void _write_delete();
    void _write_cat_val_axis();
    void _write_val_axis(val_axis_args *args);
    void _write_cat_axis(val_axis_args *args);
    void _write_chart_title();
    void _write_title(chart_title *title);
    std::vector<std::shared_ptr<chart_series> > _get_primary_axes_series();
    std::vector<std::shared_ptr<chart_series> > _get_secondary_axes_series();
    void _write_xval_ser(const std::shared_ptr<chart_series> &series);
    void _xml_declaration();
    void _write_print_settings();
    void _write_chart();
    void _write_bar_dir(const std::string &type);
};

typedef std::shared_ptr<chart> chart_ptr;

class chart_area : public chart {
public:
    chart_area(uint8_t t) : chart(t) {}
protected:
    void write_chart_type(bool);
    void _initialize();
};

class chart_bar : public chart {
public:
    chart_bar(uint8_t t) : chart(t) {}
protected:
    void write_chart_type(bool);
    void _initialize();
};

class chart_column : public chart {
public:
    chart_column(uint8_t t) : chart(t) {}
protected:
    void write_chart_type(bool);
    void _initialize();
};

class chart_line : public chart {
public:
    chart_line(uint8_t t) : chart(t) {}
protected:
    void write_chart_type(bool);
    void _initialize();
};

class chart_pie : public chart {
public:
    chart_pie(uint8_t t) : chart(t) {}
protected:
    void write_chart_type(bool);
    void write_plot_area();
    void _initialize();
};

class chart_scatter: public chart {
public:
    chart_scatter(uint8_t t) : chart(t) {}
    chart_series* add_series(const std::string& categories = std::string(), const std::string& values = std::string(), const series_options& options = series_options());
protected:
    void write_chart_type(bool);
    void write_plot_area();
    void _initialize();
};

class chart_radar : public chart {
public:
    chart_radar(uint8_t t) : chart(t) {}
protected:
    void write_chart_type(bool);
    void _initialize();
};

class chart_doughtnut : public chart_pie {
public:
    chart_doughtnut(uint8_t t) : chart_pie(t) {}
protected:
    void write_chart_type(bool);
    void _initialize();
};

int XLSXWRITER_EXPORT chart_add_data_cache(series_range *range, uint8_t *data,
                             uint16_t rows, uint8_t cols, uint8_t col);


} // namespace xlsxwriter

#endif /* __LXW_CHART_H__ */
