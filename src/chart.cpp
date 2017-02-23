/*****************************************************************************
 * chart - A library for creating Excel XLSX chart files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xmlwriter.hpp"
#include "chart.hpp"
#include "utility.hpp"
#include <math.h>

namespace xlsxwriter {

typedef struct val_axis_args {
    std::shared_ptr<chart_axis> x_axis;
    std::shared_ptr<chart_axis> y_axis;
    uint32_t id_1;
    uint32_t id_2;
} val_axis_args;

/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Free a series range object.
 */
void void_chart_free_range(series_range *range)
{
    /*
    struct lxw_series_data_point *data_point;

    if (!range)
        return;

    while (!STAILQ_EMPTY(range->data_cache)) {
        data_point = STAILQ_FIRST(range->data_cache);
        free(data_point->string);
        STAILQ_REMOVE_HEAD(range->data_cache, list_pointers);

        free(data_point);
    }

    free(range->data_cache);
    free(range->formula);
    free(range->sheetname);
    free(range);
    */
}

/*
 * Free a series object.
 */
/*
void chart_series::series_free(lxw_chart_series *series)
{
    if (!series)
        return;

    free(series->title.name);

    _chart_free_range(series->categories);
    _chart_free_range(series->values);
    _chart_free_range(series->title.range);

    free(series);
}

*/

/*
 * Initialize the data cache in a range object.
 */
int _char_init_data_cache(series_range *range)
{
    /* Initialize the series range data cache. */
    range->data_cache = calloc(1, sizeof(struct lxw_series_data_points));
    RETURN_ON_MEM_ERROR(range->data_cache, -1);
    STAILQ_INIT(range->data_cache);

    return 0;
}

/*
 * Create a new axis object
 */
/*
lxw_chart_axis *lxw_axis_new()
{
    lxw_chart_axis *axis = calloc(1, sizeof(struct lxw_chart_axis));
    GOTO_LABEL_ON_MEM_ERROR(axis, mem_error);

    axis->min_value = NAN;
    axis->max_value = NAN;
    axis->title.range = calloc(1, sizeof(lxw_series_range));
    GOTO_LABEL_ON_MEM_ERROR(axis->title.range, mem_error);

    return axis;
mem_error:
    return NULL;
}*/

/*
 * Create a new chart object.
 */
chart::chart(uint8_t type)
{

    x_axis = std::make_shared<chart_axis>();

    y_axis = std::make_shared<chart_axis>();

    x2_axis = std::make_shared<chart_axis>();

    y2_axis = std::make_shared<chart_axis>();

    title.range = calloc(1, sizeof(series_range));
    GOTO_LABEL_ON_MEM_ERROR(title.range, mem_error);

    /* Initialize the ranges in the chart titles. */
    if (_chart_init_data_cache(title.range) != LXW_NO_ERROR)
        goto mem_error;

    if (_chart_init_data_cache(x_axis->title.range) != LXW_NO_ERROR)
        goto mem_error;

    if (_chart_init_data_cache(y_axis->title.range) != LXW_NO_ERROR)
        goto mem_error;

    if (_chart_init_data_cache(x2_axis->title.range) != LXW_NO_ERROR)
        goto mem_error;

    if (_chart_init_data_cache(y2_axis->title.range) != LXW_NO_ERROR)
        goto mem_error;

    type = type;
    style_id = 2;
    hole_size = 50;

    /* Set the default axis positions. */
    cat_axis_position = LXW_CHART_BOTTOM;
    val_axis_position = LXW_CHART_LEFT;

    /* Set the default legend position */
    legend_position = LXW_CHART_RIGHT;

    lxw_strcpy(x_axis->default_num_format, "General");
    lxw_strcpy(y_axis->default_num_format, "General");
    lxw_strcpy(x2_axis->default_num_format, "General");
    lxw_strcpy(y2_axis->default_num_format, "General");

    x_axis->default_major_gridlines = false;
    y_axis->default_major_gridlines = true;

    x_axis->visible = true;
    y_axis->visible = true;
    x2_axis->visible = false;
    y2_axis->visible = true;
    
    x2_axis->default_major_gridlines = false;
    y2_axis->default_major_gridlines = false;

    x_axis->position = LXW_CHART_BOTTOM;
    y_axis->position = LXW_CHART_LEFT;
    x2_axis->position = LXW_CHART_TOP;
    y2_axis->position = LXW_CHART_RIGHT;

    series_overlap_1 = 100;

    has_horiz_cat_axis = false;
    has_horiz_val_axis = true;
}

/*
 * Free a chart object.
 */
chart::~chart()
{
}

/*
 * Add unique ids for primary or secondary axes.
 */
void chart::_add_axis_ids(bool primary)
{
    uint32_t chart_id = 50010000 + id;
    uint32_t axis_count = 1 + (axis_id_1 > 0) + (axis_id_2 > 0) + (axis_id_3 > 0) + (axis_id_4 > 0);

    uint32_t id_1 = chart_id + axis_count;
    uint32_t id_2 = id_1 + 1;

    if (primary)
    {
        axis_id_1 = id_1;
        axis_id_2 = id_2;
    }
    else
    {
        axis_id_3 = id_1;
        axis_id_4 = id_2;
    }
}

/*
 * Utility function to set a chart range.
 */
static void chart::set_range(series_range *range, const std::string& sheetname,
                 lxw_row_t first_row, lxw_col_t first_col,
                 lxw_row_t last_row, lxw_col_t last_col)
{
    std::string formula;

    /* Set the range properties. */
    range->sheetname = sheetname;
    range->first_row = first_row;
    range->first_col = first_col;
    range->last_row = last_row;
    range->last_col = last_col;

    /* Convert the range properties to a formula like: Sheet1!$A$1:$A$5. */
    lxw_rowcol_to_formula_abs(formula, sheetname,
                              first_row, first_col, last_row, last_col);

    range->formula = formula;
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
void chart::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <c:chartSpace> element.
 */
void chart::_write_chart_space()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns_c[] = LXW_SCHEMA_DRAWING "/chart";
    char xmlns_a[] = LXW_SCHEMA_DRAWING "/main";
    char xmlns_r[] = LXW_SCHEMA_OFFICEDOC "/relationships";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns:c", xmlns_c);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:a", xmlns_a);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:r", xmlns_r);

    lxw_xml_start_tag("c:chartSpace", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:lang> element.
 */
void chart::_write_lang()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "en-US");

    lxw_xml_empty_tag("c:lang", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:style> element.
 */
void chart::_write_style()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    /* Don"t write an element for the default style, 2. */
    if (style_id == 2)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", style_id);

    lxw_xml_empty_tag("c:style", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:layout> element.
 */
void chart::_write_layout()
{
    lxw_xml_empty_tag("c:layout", NULL);
}

/*
 * Write the <c:grouping> element.
 */
void chart::_write_grouping(uint8_t grouping)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (grouping == LXW_GROUPING_STANDARD)
        LXW_PUSH_ATTRIBUTES_STR("val", "standard");
    else if (grouping == LXW_GROUPING_PERCENTSTACKED)
        LXW_PUSH_ATTRIBUTES_STR("val", "percentStacked");
    else if (grouping == LXW_GROUPING_STACKED)
        LXW_PUSH_ATTRIBUTES_STR("val", "stacked");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "clustered");

    lxw_xml_empty_tag("c:grouping", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:radarStyle> element.
 */
void chart::_write_radar_style()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (type == LXW_CHART_RADAR_FILLED)
        LXW_PUSH_ATTRIBUTES_STR("val", "filled");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "marker");

    lxw_xml_empty_tag("c:radarStyle", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:varyColors> element.
 */
void chart::_write_vary_colors()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag("c:varyColors", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:firstSliceAng> element.
 */
void chart::_write_first_slice_ang()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", rotation);

    lxw_xml_empty_tag("c:firstSliceAng", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:holeSize> element.
 */
void chart::_write_hole_size()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", hole_size);

    lxw_xml_empty_tag("c:holeSize", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:t> element.
 */
void chart::_write_a_t(const std::string& name)
{
    lxw_xml_data_element("a:t", name, NULL);
}

/*
 * Write the <a:endParaRPr> element.
 */
void chart::_write_a_end_para_rpr()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("lang", "en-US");

    lxw_xml_empty_tag("a:endParaRPr", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:defRPr> element.
 */
void chart::_write_a_def_rpr()
{
    lxw_xml_empty_tag("a:defRPr", NULL);
}

/*
 * Write the <a:rPr> element.
 */
void chart::_write_a_r_pr()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char lang[] = "en-US";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("lang", lang);

    lxw_xml_empty_tag("a:rPr", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:r> element.
 */
void chart::_write_a_r(const std::string& name)
{
    lxw_xml_start_tag("a:r", NULL);

    /* Write the a:rPr element. */
    _write_a_r_pr();

    /* Write the a:t element. */
    _write_a_t(name);

    lxw_xml_end_tag("a:r");
}

/*
 * Write the <a:pPr> element.
 */
void chart::_write_a_p_pr()
{
    lxw_xml_start_tag("a:pPr", NULL);

    /* Write the a:defRPr element. */
    _write_a_def_rpr();

    lxw_xml_end_tag("a:pPr");
}

/*
 * Write the <a:pPr> element for pie chart legends.
 */
void chart::_write_a_p_pr_pie()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("rtl", "0");

    lxw_xml_start_tag("a:pPr", &attributes);

    /* Write the a:defRPr element. */
    _write_a_def_rpr();

    lxw_xml_end_tag("a:pPr");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:pPr> element.
 */
void chart::_write_a_p_pr_rich()
{
    lxw_xml_start_tag("a:pPr", NULL);

    /* Write the a:defRPr element. */
    _write_a_def_rpr();

    lxw_xml_end_tag("a:pPr");
}

/*
 * Write the <a:p> element.
 */
void chart::_write_a_p()
{
    lxw_xml_start_tag("a:p", NULL);

    /* Write the a:pPr element. */
    _write_a_p_pr();

    /* Write the a:endParaRPr element. */
    _write_a_end_para_rpr();

    lxw_xml_end_tag("a:p");
}

/*
 * Write the <a:p> element for pie chart legends.
 */
void chart::_write_a_p_pie()
{
    lxw_xml_start_tag("a:p", NULL);

    /* Write the a:pPr element. */
    _write_a_p_pr_pie();

    /* Write the a:endParaRPr element. */
    _write_a_end_para_rpr();

    lxw_xml_end_tag("a:p");
}

/*
 * Write the <a:p> element.
 */
void chart::_write_a_p_rich(const std::string& name)
{
    lxw_xml_start_tag("a:p", NULL);

    /* Write the a:pPr element. */
    _write_a_p_pr_rich();

    /* Write the a:r element. */
    _write_a_r(name);

    lxw_xml_end_tag("a:p");
}

/*
 * Write the <a:lstStyle> element.
 */
void chart::_write_a_lst_style()
{
    lxw_xml_empty_tag("a:lstStyle", NULL);
}

/*
 * Write the <a:bodyPr> element.
 */
void chart::_write_a_body_pr(lxw_chart_title *title)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (title && title->is_horizontal) {
        LXW_PUSH_ATTRIBUTES_STR("rot", "-5400000");
        LXW_PUSH_ATTRIBUTES_STR("vert", "horz");
    }

    lxw_xml_empty_tag("a:bodyPr", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:ptCount> element.
 */
void chart::_write_pt_count(uint16_t num_data_points)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", num_data_points);

    lxw_xml_empty_tag("c:ptCount", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:v> element.
 */
void chart::_write_v_num(double number)
{
    char data[LXW_ATTR_32];

    lxw_snprintf(data, LXW_ATTR_32, "%.16g", number);

    lxw_xml_data_element("c:v", data, NULL);
}

/*
 * Write the <c:v> element.
 */
void chart::_write_v_str(const std::string& str)
{
    lxw_xml_data_element("c:v", str, NULL);
}

/*
 * Write the <c:f> element.
 */
void chart::_write_f(const std::string& formula)
{
    lxw_xml_data_element("c:f", formula, NULL);
}

/*
 * Write the <c:pt> element.
 */
void chart::_write_pt(uint16_t index, lxw_series_data_point *data_point)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    /* Ignore chart points that have no data. */
    if (data_point->no_data)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("idx", index);

    lxw_xml_start_tag("c:pt", &attributes);

    if (data_point->is_string && data_point->string)
        _write_v_str(data_point->string);
    else
        _write_v_num(data_point->number);

    lxw_xml_end_tag("c:pt");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:pt> element.
 */
void chart::_write_num_pt(uint16_t index, lxw_series_data_point *data_point)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    /* Ignore chart points that have no data. */
    if (data_point->no_data)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("idx", index);

    lxw_xml_start_tag("c:pt", &attributes);

    _write_v_num(data_point->number);

    lxw_xml_end_tag("c:pt");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:formatCode> element.
 */
 void chart::_write_format_code()
{
    lxw_xml_data_element("c:formatCode", "General", NULL);
}

/*
 * Write the <c:numCache> element.
 */
void chart::_write_num_cache(series_range *range)
{
    lxw_series_data_point *data_point;
    uint16_t index = 0;

    lxw_xml_start_tag("c:numCache", NULL);

    /* Write the c:formatCode element. */
    _write_format_code();

    /* Write the c:ptCount element. */
    _write_pt_count(range->num_data_points);

    STAILQ_FOREACH(data_point, range->data_cache, list_pointers) {
        /* Write the c:pt element. */
        _write_num_pt(index, data_point);
        index++;
    }

    lxw_xml_end_tag("c:numCache");
}

/*
 * Write the <c:strCache> element.
 */
void chart::_write_str_cache(series_range *range)
{
    lxw_series_data_point *data_point;
    uint16_t index = 0;

    lxw_xml_start_tag("c:strCache", NULL);

    /* Write the c:ptCount element. */
    _write_pt_count(range->num_data_points);

    STAILQ_FOREACH(data_point, range->data_cache, list_pointers) {
        /* Write the c:pt element. */
        _write_pt(index, data_point);
        index++;
    }

    lxw_xml_end_tag("c:strCache");
}

/*
 * Write the <c:numRef> element.
 */
void chart::_write_num_ref(series_range *range)
{
    lxw_xml_start_tag("c:numRef", NULL);

    /* Write the c:f element. */
    _write_f(range->formula);

    if (!STAILQ_EMPTY(range->data_cache)) {
        /* Write the c:numCache element. */
        _write_num_cache(range);
    }

    lxw_xml_end_tag("c:numRef");
}

/*
 * Write the <c:strRef> element.
 */
void chart::_write_str_ref(series_range *range)
{
    lxw_xml_start_tag("c:strRef", NULL);

    /* Write the c:f element. */
    _write_f(range->formula);

    if (!STAILQ_EMPTY(range->data_cache)) {
        /* Write the c:strCache element. */
        _write_str_cache(range);
    }

    lxw_xml_end_tag("c:strRef");
}

/*
 * Write the cached data elements.
 */
void chart::_write_data_cache(series_range *range, bool has_string_cache)
{
    if (has_string_cache) {
        /* Write the c:strRef element. */
        _write_str_ref(range);
    }
    else {
        /* Write the c:numRef element. */
        _write_num_ref(range);
    }
}

/*
 * Write the <c:tx> element with a simple value such as for series names.
 */
void chart::_write_tx_value(const std::string& name)
{
    lxw_xml_start_tag("c:tx", NULL);

    /* Write the c:v element. */
    _write_v_str(name);

    lxw_xml_end_tag("c:tx");
}

/*
 * Write the <c:tx> element with a simple value such as for series names.
 */
void chart::_write_tx_formula(lxw_chart_title *title)
{
    lxw_xml_start_tag("c:tx", NULL);

    _write_str_ref(title->range);

    lxw_xml_end_tag("c:tx");
}

/*
 * Write the <c:txPr> element.
 */
void chart::_write_tx_pr(lxw_chart_title *title)
{
    lxw_xml_start_tag("c:txPr", NULL);

    /* Write the a:bodyPr element. */
    _write_a_body_pr(title);

    /* Write the a:lstStyle element. */
    _write_a_lst_style();

    /* Write the a:p element. */
    _write_a_p();

    lxw_xml_end_tag("c:txPr");
}

/*
 * Write the <c:txPr> element for pie chart legends.
 */
void chart::_write_tx_pr_pie(lxw_chart_title *title)
{
    lxw_xml_start_tag("c:txPr", NULL);

    /* Write the a:bodyPr element. */
    _write_a_body_pr(title);

    /* Write the a:lstStyle element. */
    _write_a_lst_style();

    /* Write the a:p element. */
    _write_a_p_pie();

    lxw_xml_end_tag("c:txPr");
}

/*
 * Write the <c:rich> element.
 */
void chart::_write_rich(lxw_chart_title *title)
{
    lxw_xml_start_tag("c:rich", NULL);

    /* Write the a:bodyPr element. */
    _write_a_body_pr(title);

    /* Write the a:lstStyle element. */
    _write_a_lst_style();

    /* Write the a:p element. */
    _write_a_p_rich(title->name);

    lxw_xml_end_tag("c:rich");
}

/*
 * Write the <c:tx> element.
 */
void chart::_write_tx_rich(lxw_chart_title *title)
{
    lxw_xml_start_tag("c:tx", NULL);

    /* Write the c:rich element. */
    _write_rich(title);

    lxw_xml_end_tag("c:tx");
}

/*
 * Write the <c:title> element for rich strings.
 */
void chart::_write_title_rich(lxw_chart_title *title)
{
    lxw_xml_start_tag("c:title", NULL);

    /* Write the c:tx element. */
    _write_tx_rich(title);

    /* Write the c:layout element. */
    _write_layout();

    lxw_xml_end_tag("c:title");
}

/*
 * Write the <c:title> element for a formula style title
 */
void chart::_write_title_formula(lxw_chart_title *title)
{
    lxw_xml_start_tag("c:title", NULL);

    /* Write the c:tx element. */
    _write_tx_formula(title);

    /* Write the c:layout element. */
    _write_layout();

    /* Write the c:txPr element. */
    _write_tx_pr(title);

    lxw_xml_end_tag("c:title");
}

/*
 * Write the <c:autoTitleDeleted> element.
 */
void chart::_write_auto_title_deleted()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");

    lxw_xml_empty_tag("c:autoTitleDeleted", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:idx> element.
 */
void chart::_write_idx(uint16_t index)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", index);

    lxw_xml_empty_tag("c:idx", &attributes);

    LXW_FREE_ATTRIBUTES();
}

void chart::_write_a_alpha(double transparency)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char str[10];

    sprintf(str, "%d", (int)((100 - transparency)*1000));
    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", str);
    lxw_xml_empty_tag("a:alpha", &attributes);
    LXW_FREE_ATTRIBUTES();
}

void chart::_write_a_srgb(lxw_color_t color, double transparency)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char rgb_str[LXW_ATTR_32];

    lxw_snprintf(rgb_str, LXW_ATTR_32, "%06X",
        color & 0xFFFFFF);
    LXW_INIT_ATTRIBUTES();

    LXW_PUSH_ATTRIBUTES_STR("val", rgb_str);
    if (transparency)
    {
        lxw_xml_start_tag("a:srgbClr", &attributes);
        _write_a_alpha(transparency);
        lxw_xml_end_tag("a:srgbClr");
    }
    else
    {
        lxw_xml_empty_tag("a:srgbClr", &attributes);
    }
    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <a:noFill> element.
 */
void chart::_write_a_no_fill()
{
    lxw_xml_empty_tag("a:noFill", NULL);
}

void chart::_write_a_solid_fill(lxw_color_t color, double transparency)
{    
    lxw_xml_start_tag("a:solidFill", NULL);
    _write_a_srgb(color, transparency);
    lxw_xml_end_tag("a:solidFill");
}

/*
 * Write the <a:ln> element.
 */
void chart::_write_a_ln(lxw_line *line)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char w[] = "28575";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("w", w);

    if (line->none)
    {
        lxw_xml_start_tag("a:ln", &attributes);
        /* Write the a:noFill element. */
        _write_a_no_fill();
    }
    else
    {
        lxw_xml_start_tag("a:ln", NULL);
        _write_a_solid_fill(line->color, line->transparency);
    }  
    lxw_xml_end_tag("a:ln");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:spPr> element.
 */
void chart::_write_sp_pr(lxw_shape_properties* properties)
{

    if (!properties->fill.defined && !properties->line.defined && !properties->pattern.defined)
        return;

    lxw_xml_start_tag("c:spPr", NULL);
    
    if (properties->fill.defined)
    {
        if (properties->fill.none)
            _write_a_no_fill();
        else
            _write_a_solid_fill(properties->fill.color, properties->fill.transparency);
    }   
    /*
    if (properties->pattern.defined)
        _chart_write_a_patt_fill(&properties->pattern);
        */

    /* Write the a:ln element. */
    if (properties->line.defined)
        _write_a_ln(&properties->line);

    lxw_xml_end_tag("c:spPr");
}

/*
 * Write the <c:order> element.
 */
void chart::_write_order(uint16_t index)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", index);

    lxw_xml_empty_tag("c:order", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:axId> element.
 */
void chart::_write_axis_id(uint32_t axis_id)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", axis_id);

    lxw_xml_empty_tag("c:axId", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:axId> element.
 */
void chart::_write_axis_ids(uint8_t primary)
{
    /*if (!axis_id_1)*/
        _add_axis_ids(primary);

    if (primary)
    {
        _write_axis_id(axis_id_1);
        _write_axis_id(axis_id_2);
    }
    else
    {
        _write_axis_id(axis_id_3);
        _write_axis_id(axis_id_4);
    }
}

/*
 * Write the series name.
 */
void chart::_write_series_name(const std::shared_ptr<chart_series>& series)
{
    if (!series->title.name.empty()) {
        /* Write the c:tx element. */
        _write_tx_value(series->title.name);
    }
    else if (!series->title.range->formula.empty()) {
        /* Write the c:tx element. */
        _write_tx_formula(&series->title);
    }
}

/*
 * Write the <c:majorTickMark> element.
 */
void chart::_write_major_tick_mark(const std::shared_ptr<chart_axis>& axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    if (!axis->major_tick_mark)
        return;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "cross");

    lxw_xml_empty_tag("c:majorTickMark", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:symbol> element.
 */
void chart::_write_symbol()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char val[] = "none";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", val);

    lxw_xml_empty_tag("c:symbol", &attributes);

    LXW_FREE_ATTRIBUTES();
}

void chart::_write_marker_data(lxw_marker *marker)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char val[] = "none";

    LXW_INIT_ATTRIBUTES();

    switch (marker->marker_type)
    {
    case LXW_MARKER_NONE:
        LXW_PUSH_ATTRIBUTES_STR("val", "none");
        break;
    case LXW_MARKER_TRIANGLE:
        LXW_PUSH_ATTRIBUTES_STR("val", "triangle");
        break;
    case LXW_MARKER_DIAMOND:
        LXW_PUSH_ATTRIBUTES_STR("val", "diamond");
        break;
    case LXW_MARKER_SQUARE:
        LXW_PUSH_ATTRIBUTES_STR("val", "square");
        break;
    default:
        LXW_PUSH_ATTRIBUTES_STR("val", val);
        break;
    }

    lxw_xml_empty_tag("c:symbol", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:marker> element.
 */
void chart::_write_marker(lxw_marker* marker)
{
    if (!has_markers)
        return;
    /*
    if (marker->marker_type == LXW_MARKER_NONE)
        return;
        */

    lxw_xml_start_tag("c:marker", NULL);

    _write_marker_data(marker);

    if (!(marker->marker_type == LXW_MARKER_NONE))
    {
        _write_sp_pr(&marker->properties);
    }
    
    /* Write the c:symbol element. */
    //_chart_write_symbol();

    lxw_xml_end_tag("c:marker");
}

/*
 * Write the <c:marker> element.
 */
void chart::_write_marker_value()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char val[] = "1";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", val);

    lxw_xml_empty_tag("c:marker", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:smooth> element.
 */
void chart::_write_smooth()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char val[] = "1";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", val);

    lxw_xml_empty_tag("c:smooth", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:scatterStyle> element.
 */
void chart::_write_scatter_style()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (type == LXW_CHART_SCATTER_SMOOTH
        || type == LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS)
        LXW_PUSH_ATTRIBUTES_STR("val", "smoothMarker");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "lineMarker");

    lxw_xml_empty_tag("c:scatterStyle", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:cat> element.
 */
void chart::_write_cat(const std::shared_ptr<chart_series>& series)
{
    bool has_string_cache = series->categories->has_string_cache;

    /* Ignore <c:cat> elements for charts without category values. */
    if (series->categories->formula.empty())
        return;

    cat_has_num_fmt = !has_string_cache;

    lxw_xml_start_tag("c:cat", NULL);

    /* Write the c:numRef element. */
    _write_data_cache(series->categories, has_string_cache);

    lxw_xml_end_tag("c:cat");
}

/*
 * Write the <c:xVal> element.
 */
void chart::_write_x_val(const std::shared_ptr<chart_series>& series)
{
    bool has_string_cache = series->categories->has_string_cache;

    lxw_xml_start_tag("c:xVal", NULL);

    /* Write the data cache elements. */
    _write_data_cache(series->categories, has_string_cache);

    lxw_xml_end_tag("c:xVal");
}

/*
 * Write the <c:val> element.
 */
void chart::_write_val(const std::shared_ptr<chart_series>& series)
{
    lxw_xml_start_tag("c:val", NULL);

    /* Write the data cache elements. The string_cache is set to false since
     * this should always be a number series. */
    _write_data_cache(series->values, false);

    lxw_xml_end_tag("c:val");
}

/*
 * Write the <c:yVal> element.
 */
void chart::_write_y_val(const std::shared_ptr<chart_series>& series)
{
    lxw_xml_start_tag("c:yVal", NULL);

    /* Write the data cache elements. The string_cache is set to false since
     * this should always be a number series. */
    _write_data_cache(series->values, false);

    lxw_xml_end_tag("c:yVal");
}

/*
 * Write the <c:ser> element.
 */
void chart::_write_ser(const std::shared_ptr<chart_series>& series)
{
    uint16_t index = series_index++;

    lxw_xml_start_tag("c:ser", NULL);

    /* Write the c:idx element. */
    _write_idx(index);

    /* Write the c:order element. */
    _write_order(index);

    /* Write the series name. */
    _write_series_name(series);

    /*if (series->marker.marker_type != LXW_MARKER_NONE)*/
        _write_sp_pr(&series->properties);

    /* Write the c:marker element. */
    _write_marker(&series->marker);

    /* Write the c:cat element. */
    _write_cat(series);

    /* Write the c:val element. */
    _write_val(series);

    lxw_xml_end_tag("c:ser");
}

/*
 * Write the <c:ser> element but with c:xVal/c:yVal instead of c:cat/c:val
 * elements.
 */
void chart::_write_xval_ser(const std::shared_ptr<chart_series>& series)
{
    uint16_t index = series_index++;

    lxw_xml_start_tag("c:ser", NULL);

    /* Write the c:idx element. */
    _write_idx(index);

    /* Write the c:order element. */
    _write_order(index);

   /* if (type == LXW_CHART_SCATTER)*/ {
        /* Write the c:spPr element. */
        _write_sp_pr(&series->properties);
    }

    if (!series->title.name.empty())
        _write_tx_value(series->title.name);

    if (type == LXW_CHART_SCATTER_STRAIGHT
        || type == LXW_CHART_SCATTER_SMOOTH) {
        /* Write the c:marker element. */
        _write_marker(&series->marker);
    }

    /* Write the c:xVal element. */
    _write_x_val(series);

    /* Write the yVal element. */
    _write_y_val(series);

    if (type == LXW_CHART_SCATTER_SMOOTH
        || type == LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS) {
        /* Write the c:smooth element. */
        _write_smooth();
    }

    lxw_xml_end_tag("c:ser");
}

/*
 * Write the <c:orientation> element.
 */
void chart::_write_orientation()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char val[] = "minMax";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", val);

    lxw_xml_empty_tag("c:orientation", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:scaling> element.
 */
void chart::_write_scaling(const std::shared_ptr<chart_axis>& axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_xml_start_tag("c:scaling", NULL);

    /* Write the c:orientation element. */
    _write_orientation();

    if (!isnan(axis->min_value)) {
        LXW_INIT_ATTRIBUTES();
        LXW_PUSH_ATTRIBUTES_DBL("val", axis->min_value);

        lxw_xml_empty_tag("c:min", &attributes);

        LXW_FREE_ATTRIBUTES();
    }
    if (!isnan(axis->max_value)) {
        LXW_INIT_ATTRIBUTES();
        LXW_PUSH_ATTRIBUTES_DBL("val", axis->max_value);

        lxw_xml_empty_tag("c:max", &attributes);

        LXW_FREE_ATTRIBUTES();
    }

    lxw_xml_end_tag("c:scaling");
}

/*
 * Write the <c:axPos> element.
 */
void chart::_write_axis_pos(uint8_t position)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (position == LXW_CHART_RIGHT)
        LXW_PUSH_ATTRIBUTES_STR("val", "r");
    else if (position == LXW_CHART_LEFT)
        LXW_PUSH_ATTRIBUTES_STR("val", "l");
    else if (position == LXW_CHART_TOP)
        LXW_PUSH_ATTRIBUTES_STR("val", "t");
    else if (position == LXW_CHART_BOTTOM)
        LXW_PUSH_ATTRIBUTES_STR("val", "b");

    lxw_xml_empty_tag("c:axPos", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:tickLblPos> element.
 */
void chart::_write_tick_lbl_pos()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char val[] = "nextTo";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", val);

    lxw_xml_empty_tag("c:tickLblPos", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:crossAx> element.
 */
void chart::_write_cross_axis(uint32_t axis_id)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", axis_id);

    lxw_xml_empty_tag("c:crossAx", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:crosses> element.
 */
void chart::_write_crosses(const std::string& value)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char default_val[] = "autoZero";

    LXW_INIT_ATTRIBUTES();
    if (value.empty())
        LXW_PUSH_ATTRIBUTES_STR("val", default_val);
    else
        LXW_PUSH_ATTRIBUTES_STR("val", value);

    lxw_xml_empty_tag("c:crosses", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:auto> element.
 */
void chart::_write_auto()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char val[] = "1";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", val);

    lxw_xml_empty_tag("c:auto", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:lblAlgn> element.
 */
void chart::_write_lbl_algn()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char val[] = "ctr";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", val);

    lxw_xml_empty_tag("c:lblAlgn", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:lblOffset> element.
 */
void chart::_write_lbl_offset()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char val[] = "100";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", val);

    lxw_xml_empty_tag("c:lblOffset", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:majorGridlines> element.
 */
void chart::_write_major_gridlines(const std::shared_ptr<chart_axis>& axis)
{

    if (axis->default_major_gridlines)
        lxw_xml_empty_tag("c:majorGridlines", NULL);
}

/*
 * Write the <c:numFmt> element.
 */
void chart::_write_number_format(const std::shared_ptr<chart_axis>& axis)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    if (!axis->num_format.empty()) {
        LXW_PUSH_ATTRIBUTES_STR("formatCode", axis->num_format);
        LXW_PUSH_ATTRIBUTES_STR("sourceLinked", "0");
    }
    else {
        LXW_PUSH_ATTRIBUTES_STR("formatCode", axis->default_num_format);
        LXW_PUSH_ATTRIBUTES_STR("sourceLinked", "1");
    }

    lxw_xml_empty_tag("c:numFmt", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:crossBetween> element.
 */
void
chart::_write_cross_between()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();

    if (cross_between)
        LXW_PUSH_ATTRIBUTES_STR("val", "midCat");
    else
        LXW_PUSH_ATTRIBUTES_STR("val", "between");

    lxw_xml_empty_tag("c:crossBetween", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:legendPos> element.
 */
void chart::_write_legend_pos()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;    

    LXW_INIT_ATTRIBUTES();
    switch (legend_position) {
    case LXW_CHART_RIGHT:
        LXW_PUSH_ATTRIBUTES_STR("val", "r");
        break;
    case LXW_CHART_LEFT:
        LXW_PUSH_ATTRIBUTES_STR("val", "l");
        break;
    case LXW_CHART_TOP:
        LXW_PUSH_ATTRIBUTES_STR("val", "t");
        break;
    case LXW_CHART_BOTTOM:
        LXW_PUSH_ATTRIBUTES_STR("val", "b");
        break;
    default:
        LXW_PUSH_ATTRIBUTES_STR("val", "r");
    }

    lxw_xml_empty_tag("c:legendPos", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:legend> element.
 */
void chart::_write_legend()
{
    lxw_xml_start_tag("c:legend", NULL);

    /* Write the c:legendPos element. */
    _write_legend_pos();

    /* Write the c:layout element. */
    _write_layout();

    if (type == LXW_CHART_PIE || type == LXW_CHART_DOUGHNUT) {
        /* Write the c:txPr element. */
        _write_tx_pr_pie(NULL);
    }

    lxw_xml_end_tag("c:legend");
}

/*
 * Write the <c:plotVisOnly> element.
 */
void chart::_write_plot_vis_only()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char val[] = "1";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", val);

    lxw_xml_empty_tag("c:plotVisOnly", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:headerFooter> element.
 */
void chart::_write_header_footer()
{
    lxw_xml_empty_tag("c:headerFooter", NULL);
}

/*
 * Write the <c:pageMargins> element.
 */
void chart::_write_page_margins()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char b[] = "0.75";
    char l[] = "0.7";
    char r[] = "0.7";
    char t[] = "0.75";
    char header[] = "0.3";
    char footer[] = "0.3";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("b", b);
    LXW_PUSH_ATTRIBUTES_STR("l", l);
    LXW_PUSH_ATTRIBUTES_STR("r", r);
    LXW_PUSH_ATTRIBUTES_STR("t", t);
    LXW_PUSH_ATTRIBUTES_STR("header", header);
    LXW_PUSH_ATTRIBUTES_STR("footer", footer);

    lxw_xml_empty_tag("c:pageMargins", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:pageSetup> element.
 */
void chart::_write_page_setup()
{
    lxw_xml_empty_tag("c:pageSetup", NULL);
}

/*
 * Write the <c:printSettings> element.
 */
void chart::_write_print_settings()
{
    lxw_xml_start_tag("c:printSettings", NULL);

    /* Write the c:headerFooter element. */
    _write_header_footer();

    /* Write the c:pageMargins element. */
    _write_page_margins();

    /* Write the c:pageSetup element. */
    _write_page_setup();

    lxw_xml_end_tag("c:printSettings");
}

/*
 * Write the <c:overlap> element.
 */
void chart::_write_overlap(int overlap)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("val", overlap);

    lxw_xml_empty_tag("c:overlap", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:title> element.
 */
void chart::_write_title(lxw_chart_title *title)
{
    if (!title->name.empty()) {
        /* Write the c:title element. */
        _write_title_rich(title);
    }
    else if (!title->range->formula.empty()) {
        /* Write the c:title element. */
        _write_title_formula(title);
    }
}

/*
 * Write the <c:title> element.
 */
void chart::_write_chart_title()
{
    if (title.off) {
        /* Write the c:autoTitleDeleted element. */
        _write_auto_title_deleted();
    }
    else {
        /* Write the c:title element. */
        _write_title(&title);
    }
}

void chart::_write_delete()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", "1");
    lxw_xml_empty_tag("c:delete", &attributes);
    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <c:catAx> element. Usually the X axis.
 */
void chart::_write_cat_axis(val_axis_args* args)
{
    if (!args->id_1 && !args->id_2)
        return;

    uint8_t position = args->x_axis->position;/* cat_axis_position;*/

    lxw_xml_start_tag("c:catAx", NULL);

    _write_axis_id(args->id_1);

    /* Write the c:scaling element. */
    _write_scaling(args->x_axis);

    /* Write the c:axPos element. */
    _write_axis_pos(position);

    if (!args->x_axis->visible)
        _write_delete();

    /* Write the c:majorGridlines element. */
    _write_major_gridlines(args->x_axis);

    /* Write the axis title elements. */
    args->x_axis->title.is_horizontal = has_horiz_cat_axis;
    _write_title(&args->x_axis->title);

    /* Write the c:numFmt element. */
    if (cat_has_num_fmt)
        _write_number_format(args->x_axis);

    /* Write the c:majorTickMark element. */
    _write_major_tick_mark(args->x_axis);

    /* Write the c:tickLblPos element. */
    _write_tick_lbl_pos();

    /* Write the c:crossAx element. */
    _write_cross_axis(args->id_2);

    /* Write the c:crosses element. */
    _write_crosses(args->x_axis->crossing);

    /* Write the c:auto element. */
    _write_auto();

    /* Write the c:lblAlgn element. */
    _write_lbl_algn();

    /* Write the c:lblOffset element. */
    _write_lbl_offset();

    lxw_xml_end_tag("c:catAx");
}

/*
 * Write the <c:valAx> element.
 */
void chart::_write_val_axis(val_axis_args* args)
{
    uint8_t position = args->y_axis->position;/*val_axis_position;*/

    if (!args->id_1 && !args->id_2)
        return;

    lxw_xml_start_tag("c:valAx", NULL);

    _write_axis_id(args->id_2);

    /* Write the c:scaling element. */
    _write_scaling(args->y_axis);

    /* Write the c:axPos element. */
    _write_axis_pos(position);

    if (!args->y_axis->visible)
        _write_delete();

    /* Write the c:majorGridlines element. */
    _write_major_gridlines(args->y_axis);

    /* Write the axis title elements. */
    args->y_axis->title.is_horizontal = has_horiz_val_axis;
    _write_title(&args->y_axis->title);

    /* Write the c:numFmt element. */
    _write_number_format(args->y_axis);

    /* Write the c:majorTickMark element. */
    _write_major_tick_mark(args->y_axis);

    /* Write the c:tickLblPos element. */
    _write_tick_lbl_pos();

    /* Write the c:crossAx element. */
    _write_cross_axis(args->id_1);

    /* Write the c:crosses element. */
    _write_crosses(args->y_axis->crossing);

    /* Write the c:crossBetween element. */
    _write_cross_between();

    lxw_xml_end_tag("c:valAx");
}

/*
 * Write the <c:valAx> element. This is for the second valAx in scatter plots.
 */
void chart::_write_cat_val_axis()
{
    uint8_t position = cat_axis_position;

    lxw_xml_start_tag("c:valAx", NULL);

    _write_axis_id(axis_id_1);

    /* Write the c:scaling element. */
    _write_scaling(x_axis);

    /* Write the c:axPos element. */
    _write_axis_pos(position);

    /* Write the axis title elements. */
    x_axis->title.is_horizontal = has_horiz_val_axis;
    _write_title(&x_axis->title);

    /* Write the c:numFmt element. */
    _write_number_format(x_axis);

    if (!x_axis->visible)
        _write_delete();

    /* Write the c:majorTickMark element. */
    _write_major_tick_mark(x_axis);

    /* Write the c:tickLblPos element. */
    _write_tick_lbl_pos();

    /* Write the c:crossAx element. */
    _write_cross_axis(axis_id_2);

    /* Write the c:crosses element. */
    _write_crosses(x_axis->crossing);

    /* Write the c:crossBetween element. */
    _write_cross_between();

    lxw_xml_end_tag("c:valAx");
}


std::vector<std::shared_ptr<chart_series>> chart::_get_secondary_axes_series()
{
    std::vector<std::shared_ptr<chart_series>> secondary_series_list;
    for(const auto& series : series_list) {
        if (series->y2_axis)
            secondary_series_list.push_back(series);
    }
    return secondary_series_list;
}

std::vector<std::shared_ptr<chart_series>> chart::_get_primary_axes_series()
{
    std::vector<std::shared_ptr<chart_series>> primary_series_list;
    for (const auto& series : series_list) {
        if (!series->y2_axis)
            primary_series_list.push_back(series);
    }
    return primary_series_list;
}

/*
 * Write the <c:barDir> element.
 */
void chart::_write_bar_dir(const std::string& type)
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("val", type);

    lxw_xml_empty_tag("c:barDir", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write a area chart.
 */
void chart::_write_area_chart(bool primary_axes)
{
    std::vector<std::shared_ptr<chart_series>> writable_series;

    if (primary_axes)
    {
        writable_series = _get_primary_axes_series();
    }
    else
    {
        writable_series = _get_secondary_axes_series();
    }
    if (writable_series.empty())
        return;

    lxw_xml_start_tag("c:areaChart", NULL);

    /* Write the c:grouping element. */
    _write_grouping(grouping);

    for(const auto& series : writable_series) {
        /* Write the c:ser element. */
        _write_ser(series);
    }

    if (has_overlap) {
        /* Write the c:overlap element. */
        _write_overlap(series_overlap_1);
    }

    /* Write the c:axId elements. */
    _write_axis_ids(primary_axes);

    lxw_xml_end_tag("c:areaChart");
}

/*
 * Write a bar chart.
 */
void chart::_write_bar_chart(bool primary_axes)
{
    std::vector<std::shared_ptr<chart_series>> writable_series;

    if (primary_axes)
    {
        writable_series = _get_primary_axes_series();
    }
    else
    {
        writable_series = _get_secondary_axes_series();
    }
    if (STAILQ_EMPTY(series_list))
        return;

    lxw_xml_start_tag("c:barChart", NULL);

    /* Write the c:barDir element. */
    _write_bar_dir("bar");

    /* Write the c:grouping element. */
    _write_grouping(grouping);

    for(const auto& series: writable_series) {
        /* Write the c:ser element. */
        _write_ser(series);
    }

    if (has_overlap) {
        /* Write the c:overlap element. */
        _write_overlap(series_overlap_1);
    }

    /* Write the c:axId elements. */
    _write_axis_ids(primary_axes);

    lxw_xml_end_tag("c:barChart");
}

/*
 * Write a column chart.
 */
void chart::_write_column_chart(bool primary_axes)
{
    std::vector<std::shared_ptr<chart_series>> writable_series;

    if (primary_axes)
    {
        writable_series = _get_primary_axes_series();
    }
    else
    {
        writable_series = _get_secondary_axes_series();
    }
    if (writable_series.empty())
        return;

    lxw_xml_start_tag("c:barChart", NULL);

    /* Write the c:barDir element. */
    _write_bar_dir("col");

    /* Write the c:grouping element. */
    _write_grouping(grouping);

    for(const auto& series : writable_series) {
        /* Write the c:ser element. */
        _write_ser(series);
    }

    if (has_overlap) {
        /* Write the c:overlap element. */
        _write_overlap(series_overlap_1);
    }

    /* Write the c:axId elements. */
    _write_axis_ids(primary_axes);

    lxw_xml_end_tag("c:barChart");
}

/*
 * Write a doughnut chart.
 */
void chart::_write_doughnut_chart(bool primary_axes)
{
    std::vector<std::shared_ptr<chart_series>> writable_series;

    if (primary_axes)
    {
        writable_series = _get_primary_axes_series();
    }
    else
    {
        writable_series = _get_secondary_axes_series();
    }
    if (writable_series.empty())
        return;

    lxw_xml_start_tag("c:doughnutChart", NULL);

    /* Write the c:varyColors element. */
    _write_vary_colors();

    for(const auto& series: writable_series) {
        /* Write the c:ser element. */
        _write_ser(series);
    }

    /* Write the c:firstSliceAng element. */
    _write_first_slice_ang();

    /* Write the c:holeSize element. */
    _write_hole_size();

    lxw_xml_end_tag("c:doughnutChart");
}

/*
 * Write a line chart.
 */
void chart::_write_line_chart(bool primary_axes)
{
    std::vector<std::shared_ptr<chart_series>> writable_series;

    if (primary_axes)
    {
        writable_series = _get_primary_axes_series();
    }
    else
    {
        writable_series = _get_secondary_axes_series();
    }
    if (writable_series.empty())
    {
        return;
    }

    lxw_xml_start_tag("c:lineChart", NULL);

    /* Write the c:grouping element. */
    _write_grouping(grouping);

    for(const auto& series: writable_series) {
        /* Write the c:ser element. */
        _write_ser(series);
    }
    
    /*
    lxw_marker marker = {0};
     Write the c:marker element. 
    _chart_write_marker(&marker);
    */

    /* Write the c:axId elements. */
    _write_axis_ids(primary_axes);

    lxw_xml_end_tag("c:lineChart");
}

/*
 * Write a pie chart.
 */
void chart::_write_pie_chart(uint8_t primary_axes)
{
    std::vector<std::shared_ptr<chart_series>> writable_series;

    if (primary_axes)
    {
        writable_series = _get_primary_axes_series();
    }
    else
    {
        writable_series = _get_secondary_axes_series();
    }
    if (writable_series.empty())
    {
        return;
    }

    lxw_xml_start_tag("c:pieChart", NULL);

    /* Write the c:varyColors element. */
    _write_vary_colors();

    for(const auto& series: writable_series) {
        /* Write the c:ser element. */
        _write_ser(series);
    }

    /* Write the c:firstSliceAng element. */
    _write_first_slice_ang();

    lxw_xml_end_tag("c:pieChart");
}

/*
 * Write a scatter chart.
 */
void chart::_write_scatter_chart(bool primary_axes)
{
    std::vector<std::shared_ptr<chart_series>> writable_series;

    if (primary_axes)
    {
        writable_series = _get_primary_axes_series();
    }
    else
    {
        writable_series = _get_secondary_axes_series();
    }

    if (writable_series.empty())
    {
        return;
    }

    lxw_xml_start_tag("c:scatterChart", NULL);

    /* Write the c:scatterStyle element. */
    _write_scatter_style();

    for( const auto& series : writable_series) {
        /* Write the c:ser element. */
        _write_xval_ser(series);
    }

    /* Write the c:axId elements. */
    _write_axis_ids(primary_axes);

    lxw_xml_end_tag("c:scatterChart");
}

/*
 * Write a radar chart.
 */
void chart::_write_radar_chart(bool primary_axes)
{
    std::vector<std::shared_ptr<chart_series>> series_list;

    if (primary_axes)
    {
        series_list = _get_primary_axes_series();
    }
    else
    {
        series_list = _get_secondary_axes_series();
    }
    if (series_list.empty())
    {
        return;
    }

    lxw_xml_start_tag("c:radarChart", NULL);

    /* Write the c:radarStyle element. */
    _write_radar_style();

    for(const auto& series : series_list)
    {
        /* Write the c:ser element. */
        _write_ser(series);
    }

    if (has_overlap) {
        /* Write the c:overlap element. */
        _write_overlap(series_overlap_1);
    }

    /* Write the c:axId elements. */
    _write_axis_ids(primary_axes);

    lxw_xml_end_tag("c:radarChart");
}

/*
 * Write the <c:plotArea> element.
 */
void chart::_write_scatter_plot_area()
{
    lxw_xml_start_tag("c:plotArea", NULL);

    /* Write the c:layout element. */
    _write_layout();

    /* Write subclass chart type elements for primary and secondary axes. */
    write_chart_type(this, true);
    write_chart_type(this, false);

    /* Write the c:catAx element. */
    _write_cat_val_axis();

    has_horiz_val_axis = true;

    val_axis_args args;
    args.x_axis = x_axis;
    args.y_axis = y_axis;
    args.id_1 = axis_id_1;
    args.id_2 = axis_id_2;

    /* Write the c:valAx element. */
    _write_val_axis(&args);

    lxw_xml_end_tag("c:plotArea");
}

/*
 * Write the <c:plotArea> element. Special handling for pie/doughnut.
 */
void chart::_write_pie_plot_area()
{
    lxw_xml_start_tag("c:plotArea", NULL);

    /* Write the c:layout element. */
    _write_layout();

    /* Write subclass chart type elements for primary and secondary axes. */
    write_chart_type(this, true);

    lxw_xml_end_tag("c:plotArea");
}

/*
 * Write the <c:plotArea> element.
 */
void chart::_write_plot_area()
{
    lxw_xml_start_tag("c:plotArea", NULL);

    /* Write the c:layout element. */
    _write_layout();

    /* Write subclass chart type elements for primary and secondary axes. */
    write_chart_type(this, true);
    write_chart_type(this, false);

    /* Write combined chart, if exist*/
    
    const std::shared_ptr<chart>& second_chart = combined;
    if (second_chart)
    {
        _chart_initialize(second_chart, second_chart->type);

        second_chart->id = second_chart->is_secondary ? id + 1000 : id;
        second_chart->file = file;
        second_chart->series_index = series_index;
        second_chart->write_chart_type(combined.get(), true);
        second_chart->write_chart_type(combined.get(), false);
    }
    
    

    val_axis_args args;
    args.x_axis = x_axis;
    args.y_axis = y_axis;
    args.id_1 = axis_id_1;
    args.id_2 = axis_id_2;

    /* Write the c:catAx element. */
    _write_cat_axis(&args);

    /* Write the c:valAx element. */
    _write_val_axis(&args);

    args.x_axis = x2_axis;
    args.y_axis = y2_axis;
    args.id_1 = axis_id_3;
    args.id_2 = axis_id_4;

    /* Write the c:valAx element. */
    _write_val_axis(&args);

    if (second_chart && second_chart->is_secondary)
    {
        args.x_axis = second_chart->x2_axis;
        args.y_axis = second_chart->y2_axis;
        args.id_1 = second_chart->axis_id_3;
        args.id_2 = second_chart->axis_id_4;
        _chart_write_val_axis(second_chart, &args);
    }

    _write_cat_axis(&args);

    /* TODO add c:dTable elemnt */
    /* TODO add c:spPr element */

    lxw_xml_end_tag("c:plotArea");
}

/*
 * Write the <c:chart> element.
 */
void chart::_write_chart()
{
    lxw_xml_start_tag("c:chart", NULL);

    /* Write the c:title element. */
    _write_chart_title();

    /* Write the c:plotArea element. */
    write_plot_area(this);

    /* Write the c:legend element. */
    _write_legend();

    /* Write the c:plotVisOnly element. */
    _write_plot_vis_only();

    lxw_xml_end_tag("c:chart");
}

/*
 * Initialize a area chart.
 */
void chart::_initialize_area_chart(uint8_t type)
{
    grouping = LXW_GROUPING_STANDARD;
    cross_between = LXW_CHART_AXIS_POSITION_ON_TICK;

    if (type == LXW_CHART_AREA_STACKED) {
        grouping = LXW_GROUPING_STACKED;
        subtype = LXW_CHART_SUBTYPE_STACKED;
    }

    if (type == LXW_CHART_AREA_STACKED_PERCENT) {
        grouping = LXW_GROUPING_PERCENTSTACKED;
        lxw_strcpy((y_axis)->default_num_format, "0%");
        subtype = LXW_CHART_SUBTYPE_STACKED;
    }

    /* Initialize the function pointers for this chart type. */
    write_chart_type = _chart_write_area_chart;
    write_plot_area = _chart_write_plot_area;
}

/*
 * Initialize a bar chart.
 */
void chart::_initialize_bar_chart(uint8_t type)
{
    std::shared_ptr<chart_axis> tmp;

    /* Reverse the X and Y axes for Bar charts. */
    tmp = x_axis;
    x_axis = y_axis;
    y_axis = tmp;

    /*Also reverse some of the defaults. */
    x_axis->default_major_gridlines = false;
    y_axis->default_major_gridlines = true;
    has_horiz_cat_axis = true;
    has_horiz_val_axis = false;

    if (type == LXW_CHART_BAR_STACKED) {
        grouping = LXW_GROUPING_STACKED;
        has_overlap = true;
        subtype = LXW_CHART_SUBTYPE_STACKED;
    }

    if (type == LXW_CHART_BAR_STACKED_PERCENT) {
        grouping = LXW_GROUPING_PERCENTSTACKED;
        lxw_strcpy((y_axis)->default_num_format, "0%");
        has_overlap = true;
        subtype = LXW_CHART_SUBTYPE_STACKED;
    }

    /* Override the default axis positions for a bar chart. */
    cat_axis_position = LXW_CHART_LEFT;
    val_axis_position = LXW_CHART_BOTTOM;

    /* Initialize the function pointers for this chart type. */
    write_chart_type = _chart_write_bar_chart;
    write_plot_area = _chart_write_plot_area;
}

/*
 * Initialize a column chart.
 */
void chart::_initialize_column_chart(uint8_t type)
{
    has_horiz_val_axis = false;

    if (type == LXW_CHART_COLUMN_STACKED) {
        grouping = LXW_GROUPING_STACKED;
        has_overlap = true;
        subtype = LXW_CHART_SUBTYPE_STACKED;
    }

    if (type == LXW_CHART_COLUMN_STACKED_PERCENT) {
        grouping = LXW_GROUPING_PERCENTSTACKED;
        lxw_strcpy((y_axis)->default_num_format, "0%");
        has_overlap = true;
        subtype = LXW_CHART_SUBTYPE_STACKED;
    }

    /* Initialize the function pointers for this chart type. */
    write_chart_type = _chart_write_column_chart;
    write_plot_area = _chart_write_plot_area;
}

/*
 * Initialize a doughnut chart.
 */
void chart::_initialize_doughnut_chart()
{
    has_markers = false;

    /* Initialize the function pointers for this chart type. */
    write_chart_type = _chart_write_doughnut_chart;
    write_plot_area = _chart_write_pie_plot_area;
}

/*
 * Initialize a line chart.
 */
void chart::_initialize_line_chart()
{
    has_markers = true;
    grouping = LXW_GROUPING_STANDARD;

    /* Initialize the function pointers for this chart type. */
    write_chart_type = _chart_write_line_chart;
    write_plot_area = _chart_write_plot_area;
}

/*
 * Initialize a pie chart.
 */
void chart::_initialize_pie_chart()
{
    has_markers = false;

    /* Initialize the function pointers for this chart type. */
    write_chart_type = _chart_write_pie_chart;
    write_plot_area = _chart_write_pie_plot_area;
}

/*
 * Initialize a scatter chart.
 */
void chart::_initialize_scatter_chart()
{
    has_horiz_val_axis = false;
    cross_between = LXW_CHART_AXIS_POSITION_ON_TICK;
    is_scatter = true;
    has_markers = true;

    /* Initialize the function pointers for this chart type. */
    write_chart_type = _chart_write_scatter_chart;
    write_plot_area = _chart_write_scatter_plot_area;
}

/*
 * Initialize a radar chart.
 */
void chart::_initialize_radar_chart(uint8_t type)
{
    if (type == LXW_CHART_RADAR)
        has_markers = true;

    x_axis->default_major_gridlines = true;
    y_axis->major_tick_mark = true;

    /* Initialize the function pointers for this chart type. */
    write_chart_type = _chart_write_radar_chart;
    write_plot_area = _chart_write_plot_area;
}

/*
 * Initialize the chart specific properties.
 */
void chart::_initialize(uint8_t type)
{
    switch (type) {

        case LXW_CHART_AREA:
        case LXW_CHART_AREA_STACKED:
        case LXW_CHART_AREA_STACKED_PERCENT:
            _chart_initialize_area_chart(type);
            break;

        case LXW_CHART_BAR:
        case LXW_CHART_BAR_STACKED:
        case LXW_CHART_BAR_STACKED_PERCENT:
            _chart_initialize_bar_chart(type);
            break;

        case LXW_CHART_COLUMN:
        case LXW_CHART_COLUMN_STACKED:
        case LXW_CHART_COLUMN_STACKED_PERCENT:
            _chart_initialize_column_chart(type);
            break;

        case LXW_CHART_DOUGHNUT:
            _chart_initialize_doughnut_chart();
            break;

        case LXW_CHART_LINE:
            _chart_initialize_line_chart();
            break;

        case LXW_CHART_PIE:
            _chart_initialize_pie_chart();
            break;

        case LXW_CHART_SCATTER:
        case LXW_CHART_SCATTER_STRAIGHT:
        case LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS:
        case LXW_CHART_SCATTER_SMOOTH:
        case LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS:
            _chart_initialize_scatter_chart();
            break;

        case LXW_CHART_RADAR:
        case LXW_CHART_RADAR_WITH_MARKERS:
        case LXW_CHART_RADAR_FILLED:
            _chart_initialize_radar_chart(type);
            break;

        default:
            LXW_WARN_FORMAT1("workbook_add_chart(): "
                             "unhandled chart type '%d'", type);
    }
}

/*
 * Assemble and write the XML file.
 */
void chart::assemble_xml_file()
{
    /* Initialize the chart specific properties. */
    _initialize(type);

    /* Write the XML declaration. */
    _xml_declaration();

    /* Write the c:chartSpace element. */
    _write_chart_space();

    /* Write the c:lang element. */
    _write_lang();

    /* Write the c:style element. */
    _write_style();

    /* Write the c:chart element. */
    _write_chart();

    /* Write the c:printSettings element. */
    _write_print_settings();

    lxw_xml_end_tag("c:chartSpace");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Add data to a data cache in a range object, for testing only.
 */
int lxw_chart_add_data_cache(series_range *range, uint8_t *data,
                         uint16_t rows, uint8_t cols, uint8_t col)
{
    struct lxw_series_data_point *data_point;
    uint16_t i;

    range->ignore_cache = true;
    range->num_data_points = rows;

    /* Initialize the series range data cache. */
    for (i = 0; i < rows; i++) {
        data_point = calloc(1, sizeof(struct lxw_series_data_point));
        STAILQ_INSERT_TAIL(range->data_cache, data_point, list_pointers);
        data_point->number = data[i * cols + col];
    }

    return 0;
}

void chart::set_y2_axis(const std::shared_ptr<chart_axis>& axis)
{
    y2_axis = axis;
}

/*
 * Insert an image into the worksheet.
 */
lxw_chart_series *
chart_add_series(const char *categories, const char *values)
{

mem_error:
    _chart_series_free(series);
    return NULL;
}

std::shared_ptr<chart_series> chart::add_series(const std::string& categories, const std::string&  values, const series_options& options)
{
    std::shared_ptr<chart_series> series = new std::make_shared<chart_series>();

    /* Create a new object to hold the series. */
    series = calloc(1, sizeof(lxw_chart_series));
    GOTO_LABEL_ON_MEM_ERROR(series, mem_error);

    series->categories = calloc(1, sizeof(lxw_series_range));
    GOTO_LABEL_ON_MEM_ERROR(series->categories, mem_error);

    series->values = calloc(1, sizeof(lxw_series_range));
    GOTO_LABEL_ON_MEM_ERROR(series->values, mem_error);

    series->title.range = calloc(1, sizeof(lxw_series_range));
    GOTO_LABEL_ON_MEM_ERROR(series->title.range, mem_error);

    series->marker.marker_type = LXW_MARKER_NONE;

    if (!categories.empty()) {
        if (categories[0] == '=')
            series->categories->formula = lxw_strdup(categories + 1);
        else
            series->categories->formula = lxw_strdup(categories);
    }

    if (!values.empty()) {
        if (values[0] == '=')
            series->values->formula = lxw_strdup(values + 1);
        else
            series->values->formula = lxw_strdup(values);
    }

    if (_chart_init_data_cache(series->categories) != LXW_NO_ERROR)
        goto mem_error;

    if (_chart_init_data_cache(series->values) != LXW_NO_ERROR)
        goto mem_error;

    if (_chart_init_data_cache(series->title.range) != LXW_NO_ERROR)
        goto mem_error;

    series_list.push_back(series);

    if (options.x2_axis || options.y2_axis)
    {
        series->x2_axis = options.x2_axis;
        series->y2_axis = options.y2_axis;
        is_secondary = true;
    }
    return series;
}

/*
 * Set on of the 48 built-in Excel chart styles.
 */
void
chart_set_style(uint8_t style_id)
{
    /* The default style is 2. The range is 1 - 48 */
    if (style_id < 1 || style_id > 48)
        style_id = 2;

    style_id = style_id;
}

/*
 * Set a user defined name for a series.
 */
void
chart_series_set_name(lxw_chart_series *series, const char *name)
{
    if (!name)
        return;

    if (name[0] == '=')
        series->title.range->formula = lxw_strdup(name + 1);
    else
        series->title.name = lxw_strdup(name);
}

/*
 * Set an axis caption, with a range instead or a formula..
 */
void
chart_series_set_name_range(lxw_chart_series *series, const char *sheetname,
                            lxw_row_t row, lxw_col_t col)
{
    if (!sheetname) {
        LXW_WARN("chart_series_set_name_range(): "
                 "sheetname must be specified");
        return;
    }

    /* Start and end row, col are the same for single cell range. */
    _chart_set_range(series->title.range, sheetname, row, col, row, col);
}

/*
 * Set the categories range for a series.
 */
void
chart_series_set_categories(lxw_chart_series *series, const char *sheetname,
                            lxw_row_t first_row, lxw_col_t first_col,
                            lxw_row_t last_row, lxw_col_t last_col)
{
    if (!sheetname) {
        LXW_WARN("chart_series_set_categories(): "
                 "sheetname must be specified");
        return;
    }

    _chart_set_range(series->categories, sheetname,
                     first_row, first_col, last_row, last_col);
}

/*
 * Set the values range for a series.
 */
void
chart_series_set_values(lxw_chart_series *series, const char *sheetname,
                        lxw_row_t first_row, lxw_col_t first_col,
                        lxw_row_t last_row, lxw_col_t last_col)
{
    if (!sheetname) {
        LXW_WARN("chart_series_set_values(): sheetname must be specified");
        return;
    }

    _chart_set_range(series->values, sheetname,
                     first_row, first_col, last_row, last_col);
}

/*
 * Set an axis caption.
 */
void chart_axis::set_name(const std::string& name)
{
    if (name.empty())
        return;

    if (name[0] == '=')
        title.range->formula = lxw_strdup(name + 1);
    else
        title.name = name;
}

/*
 * Set an axis caption, with a range instead or a formula..
 */
void chart_axis::set_name_range(const std::string& sheetname, lxw_row_t row, lxw_col_t col)
{
    if (sheetname.empty()) {
        LXW_WARN("chart_axis_set_name_range(): sheetname must be specified");
        return;
    }

    /* Start and end row, col are the same for single cell range. */
    chart::set_range(title.range, sheetname, row, col, row, col);
}

void  chart_axis::set_format(const std::string& format)
{	
    if (format.empty())
		return;
    num_format = format;
}

void chart_axis::set_crossing(const std::string& crossing_str)
{
    if (crossing_str.empty())
        return;
    crossing = crossing_str;
}

/*
 * Set the chart title.
 */
void chart::title_set_name(const std::string& name)
{
    if (name.empty())
        return;

    if (name[0] == '=')
        title.range->formula = name.substr(1);
    else
        title.name = name;
}

/*
 * Set the chart title, with a range instead or a formula.
 */
void chart::title_set_name_range(const std::string& sheetname,
                           lxw_row_t row, lxw_col_t col)
{
    if (sheetname.empty()) {
        LXW_WARN("chart_title_set_name_range(): sheetname must be specified");
        return;
    }

    /* Start and end row, col are the same for single cell range. */
    set_range(title.range, sheetname, row, col, row, col);
}

/*
 * Turn off the chart title.
 */
void chart::title_off()
{
    title.off = true;
}

/*
 * Set the Pie/Doughnut chart rotation: the angle of the first slice.
 */
void chart::set_rotation(uint16_t rotation)
{
    if (rotation <= 360)
        this->rotation = rotation;
    else
        LXW_WARN_FORMAT1("chart_set_rotation(): Chart rotation '%d' outside "
                         "range: 0 <= rotation <= 360", rotation);
}

/*
 * Set the Doughnut chart hole size.
 */
void chart::set_hole_size(uint8_t size)
{
    if (size >= 10 && size <= 90)
        hole_size = size;
    else
        LXW_WARN_FORMAT1("chart_set_hole_size(): Hole size '%d' outside "
                         "Excel range: 10 <= size <= 90", size);
}

} //namespace xlsxwriter
