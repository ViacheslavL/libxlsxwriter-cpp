/*****************************************************************************
 * workbook - A library for creating Excel XLSX workbook files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <xlsxwriter/xmlwriter.hpp>
#include <xlsxwriter/workbook.hpp>
#include <xlsxwriter/utility.hpp>
#include <xlsxwriter/packager.hpp>
#include <xlsxwriter/hash_table.hpp>
#include <iostream>
#include <sstream>


namespace xlsxwriter {

/*
 * Set the default index for each format. This is only used for testing.
 */
void workbook::set_default_xf_indices()
{
    for (const auto& format : formats) {
        format->get_xf_index();
    }
}

/*
 * Iterate through the XF Format objects and give them an index to non-default
 * font elements.
 */
void workbook::_prepare_fonts()
{

    hash_table<lxw_font_ptr, uint16_t> fonts;
    uint16_t index = 0;

    for (const auto& it : used_xf_formats.order_list) {
        auto format = it.first;
        lxw_font_ptr key(format->get_font_key());

        if (key) {
            /* Look up the format in the hash table. */
            auto result = fonts.exists(key);

            if (result.second) {
                /* Font has already been used. */
                format->font_index = result.first.second;
                format->has_font = false;
                delete key;
            }
            else {
                /* This is a new font. */
                uint16_t font_index = index;
                format->font_index = index;
                format->has_font = true;
                fonts.insert(key, font_index);
                index++;
            }
        }
    }
    font_count = index;
}

/*
 * Iterate through the XF Format objects and give them an index to non-default
 * border elements.
 */
void workbook::_prepare_borders()
{
    hash_table<lxw_border_ptr, uint16_t> borders;
    uint16_t index = 0;

    for (const auto& it : used_xf_formats.order_list) {
        auto format = it.first;
        lxw_border_ptr key(format->get_border_key());

        if (key) {
            /* Look up the format in the hash table. */
            auto result = borders.exists(key);

            if (result.second) {
                /* Border has already been used. */
                format->border_index = result.first.second;
                format->has_border = false;
                delete key;
            }
            else {
                /* This is a new border. */
                uint16_t border_index = index;
                format->border_index = index;
                format->has_border = true;
                borders.insert(key, border_index);
                index++;
            }
        }
    }

    border_count = index;
}

/*
 * Iterate through the XF Format objects and give them an index to non-default
 * fill elements.
 */
void workbook::_prepare_fills()
{

    hash_table<lxw_fill_ptr, uint16_t> fills;

    uint16_t index = 2;
    lxw_fill_ptr default_fill_1 = new lxw_fill();
    lxw_fill_ptr default_fill_2 = new lxw_fill();
    uint16_t fill_index1 = 0;
    uint16_t fill_index2 = 1;

    /* Add the default fills. */
    default_fill_1->pattern = LXW_PATTERN_NONE;
    default_fill_1->fg_color = LXW_COLOR_UNSET;
    default_fill_1->bg_color = LXW_COLOR_UNSET;
    fill_index1 = 0;
    fills.insert(default_fill_1, fill_index1);

    default_fill_2->pattern = LXW_PATTERN_GRAY_125;
    default_fill_2->fg_color = LXW_COLOR_UNSET;
    default_fill_2->bg_color = LXW_COLOR_UNSET;
    fill_index2 = 1;
    fills.insert(default_fill_2, fill_index2);

    for (const auto& it : used_xf_formats.order_list) {
        auto format = it.first;
        lxw_fill_ptr key(format->get_fill_key());

        /* The following logical statements jointly take care of special */
        /* cases in relation to cell colors and patterns:                */
        /* 1. For a solid fill (pattern == 1) Excel reverses the role of */
        /*    foreground and background colors, and                      */
        /* 2. If the user specifies a foreground or background color     */
        /*    without a pattern they probably wanted a solid fill, so    */
        /*    we fill in the defaults.                                   */
        if (format->pattern == LXW_PATTERN_SOLID
            && format->bg_color != LXW_COLOR_UNSET
            && format->fg_color != LXW_COLOR_UNSET) {
            lxw_color_t tmp = format->fg_color;
            format->fg_color = format->bg_color;
            format->bg_color = tmp;
        }

        if (format->pattern <= LXW_PATTERN_SOLID
            && format->bg_color != LXW_COLOR_UNSET
            && format->fg_color == LXW_COLOR_UNSET) {
            format->fg_color = format->bg_color;
            format->bg_color = LXW_COLOR_UNSET;
            format->pattern = LXW_PATTERN_SOLID;
        }

        if (format->pattern <= LXW_PATTERN_SOLID
            && format->bg_color == LXW_COLOR_UNSET
            && format->fg_color != LXW_COLOR_UNSET) {
            format->bg_color = LXW_COLOR_UNSET;
            format->pattern = LXW_PATTERN_SOLID;
        }

        if (key) {
            /* Look up the format in the hash table. */
            auto result = fills.exists(key);

            if (result.second) {
                /* Fill has already been used. */
                format->fill_index = result.first.second;
                format->has_fill = false;
                delete key;
            }
            else {
                /* This is a new fill. */
                uint16_t fill_index = index;
                format->fill_index = index;
                format->has_fill = true;
                fills.insert(key, fill_index);
                index++;
            }
        }
    }
    fill_count = index;
}

/*
 * Iterate through the XF Format objects and give them an index to non-default
 * number format elements. Note, user defined records start from index 0xA4.
 */
void workbook::_prepare_num_formats()
{
    hash_table<std::string, uint16_t> num_formats;
    uint16_t index = 0xA4;
    uint16_t num_format_count = 0;
    uint16_t num_format_index;

    for (const auto& it : used_xf_formats.order_list) {
        auto format = it.first;
        /* Format already has a number format index. */
        if (format->num_format_index)
            continue;

        /* Check if there is a user defined number format string. */
        const std::string& num_format = format->num_format;

        if (!num_format.empty()) {
            /* Look up the num_format in the hash table. */
            auto result = num_formats.exists(num_format);

            if (!result.second) {
                /* Num_Format has already been used. */
                format->num_format_index = result.first.second;
            }
            else {
                /* This is a new num_format. */
                num_format_index = index;
                format->num_format_index = index;
                num_formats.insert(num_format, num_format_index);
                index++;
                num_format_count++;
            }
        }
    }

    this->num_format_count = num_format_count;
}

/*
 * Prepare workbook and sub-objects for writing.
 */
void workbook::_prepare_workbook()
{
    /* Set the font index for the format objects. */
    _prepare_fonts();

    /* Set the number format index for the format objects. */
    _prepare_num_formats();

    /* Set the border index for the format objects. */
    _prepare_borders();

    /* Set the fill index for the format objects. */
    _prepare_fills();

}

/*
 * Process and store the defined names. The defined names are stored with
 * the Workbook.xml but also with the App.xml if they refer to a sheet
 * range like "Sheet1!:A1". The defined names are store in sorted
 * order for consistency with Excel. The names need to be normalized before
 * sorting.
 */
lxw_error workbook::_store_defined_name(
        const std::string& name,
        const std::string& app_name,
        const std::string& formula,
        int16_t index,
        bool hidden)
{
    defined_name_ptr defined_name;
    std::string name_copy; //[LXW_DEFINED_NAME_LENGTH];
    std::string tmp_str;
    std::string worksheet_name;

    /* Do some checks on the input data */
    if (name.empty() || formula.empty())
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    if (name.size() > LXW_DEFINED_NAME_LENGTH ||
        formula.size() > LXW_DEFINED_NAME_LENGTH) {
        return LXW_ERROR_128_STRING_LENGTH_EXCEEDED;
    }

    /* Allocate a new defined_name to be added to the linked list of names. */
    defined_name = std::make_shared<xlsxwriter::defined_name>();

    /* Copy the user input string. */
    name_copy = name;

    /* Set the worksheet index or -1 for a global defined name. */
    defined_name->index = index;
    defined_name->hidden = hidden;

    /* Check for local defined names like like "Sheet1!name". */
    size_t idx = name_copy.find('!');

    if (idx < name_copy.size()) {
        tmp_str = name_copy.substr(idx + 1);
    }

    if (tmp_str.empty()) {
        /* The name is global. We just store the defined name string. */
        defined_name->name = name_copy;
    }
    else {
        /* The name is worksheet local. We need to extract the sheet name
         * and map it to a sheet index. */

        /* Split the into the worksheet name and defined name. */
        worksheet_name = name_copy;

        /* Remove any worksheet quoting. */
        if (worksheet_name[0] == '\'')
            worksheet_name = worksheet_name.substr(1);
        if (worksheet_name[worksheet_name.size() - 1] == '\'')
            worksheet_name.pop_back();

        /* Search for worksheet name to get the equivalent worksheet index. */
        for (const auto& worksheet : worksheets) {
            if ( worksheet_name == worksheet->name) {
                defined_name->index = worksheet->index;
                defined_name->normalised_sheetname = worksheet_name;
            }
        }

        /* If we didn't find the worksheet name we exit. */
        if (defined_name->index == -1)
            return LXW_ERROR_MEMORY_MALLOC_FAILED;

        defined_name->name = tmp_str;
    }

    /* Print titles and repeat title pass in the name used for App.xml. */
    if (!app_name.empty()) {
        defined_name->app_name = app_name;
        defined_name->normalised_sheetname = app_name;
    }
    else {
        defined_name->app_name = name;
    }

    /* We need to normalize the defined names for sorting. This involves
     * removing any _xlnm namespace  and converting it to lowercase. */
    idx = name_copy.find("_xlnm.");

    tmp_str = idx < name_copy.size() ? name_copy.substr(idx) : "";

    if (!tmp_str.empty())
        defined_name->normalised_name = defined_name->name.substr(6);
    else
        defined_name->normalised_name = defined_name->name;

    lxw_str_tolower(defined_name->normalised_name);
    lxw_str_tolower(defined_name->normalised_sheetname);

    /* Strip leading "=" from the formula. */
    if (formula[0] == '=')
        defined_name->formula = formula.substr(1);
    else
        defined_name->formula = formula;

    defined_names.insert(defined_name);
    return LXW_NO_ERROR;
}

/*
 * Populate the data cache of a chart data series by reading the data from the
 * relevant worksheet and adding it to the cached in the range object as a
 * list of points.
 *
 * Note, the data cache isn't strictly required by Excel but it helps if the
 * chart is embedded in another application such as PowerPoint and it also
 * helps with comparison testing.
 */
void workbook::_populate_range_data_cache(const series_range_ptr& range)
{
    uint16_t num_data_points = 0;

    /* If ignore_cache is set then don't try to populate the cache. This flag
     * may be set manually, for testing, or due to a case where the cache
     * can't be calculated.
     */
    if (range->ignore_cache)
        return;

    /* Currently we only handle 2D ranges so ensure either the rows or cols
     * are the same.
     */
    if (range->first_row != range->last_row
        && range->first_col != range->last_col) {
        range->ignore_cache = true;
        return;
    }

    /* Check that the sheetname exists. */
    xlsxwriter::worksheet* worksheet = get_worksheet_by_name(range->sheetname);
    if (!worksheet) {
        LXW_WARN_FORMAT2("workbook_add_chart(): worksheet name '%s' "
                         "in chart formula '%s' doesn't exist.",
                         range->sheetname.c_str(), range->formula.c_str());
        range->ignore_cache = true;
        return;
    }

    /* We can't read the data when worksheet optimization is on. */
    if (worksheet->optimize) {
        range->ignore_cache = true;
        return;
    }

    /* Iterate through the worksheet data and populate the range cache. */
    for (lxw_row_t row_num = range->first_row; row_num <= range->last_row; row_num++) {
        lxw_row *row_obj = worksheet->find_row(row_num);

        for (lxw_col_t col_num = range->first_col; col_num <= range->last_col;
             col_num++) {

            std::shared_ptr<series_data_point> data_point = std::make_shared<series_data_point>();
            if (!data_point) {
                range->ignore_cache = true;
                return;
            }

            lxw_cell *cell_obj = worksheet->find_cell(row_obj, col_num);

            if (cell_obj) {
                if (cell_obj->type == NUMBER_CELL) {
                    data_point->number = cell_obj->u.number;
                }

                if (cell_obj->type == STRING_CELL) {
                    data_point->string = cell_obj->sst_string;
                    data_point->is_string = true;
                    range->has_string_cache = true;
                }
            }
            else {
                data_point->no_data = true;
            }

            range->data_cache.push_back(data_point);
            num_data_points++;
        }
    }

    range->num_data_points = num_data_points;

}

/* Convert a chart range such as Sheet1!$A$1:$A$5 to a sheet name and row-col
 * dimensions, or vice-versa. This gives us the dimensions to read data back
 * from the worksheet.
 */
void workbook::_populate_range_dimensions(const series_range_ptr& range)
{

    char formula[LXW_MAX_FORMULA_RANGE_LENGTH] = { 0 };
    char *tmp_str;
    char *sheetname;

    /* If neither the range formula or sheetname is defined then this probably
     * isn't a valid range.
     */
    if (range->formula.empty() && range->sheetname.empty()) {
        range->ignore_cache = true;
        return;
    }

    /* If the sheetname is already defined it was already set via
     * chart_series_set_categories() or  series->set_values().
     */
    if (!range->sheetname.empty())
        return;

    /* Ignore non-contiguous range like (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5) */
    if (range->formula[0] == '(') {
        range->ignore_cache = true;
        return;
    }

    /* Create a copy of the formula to modify and parse into parts. */
    lxw_snprintf(formula, LXW_MAX_FORMULA_RANGE_LENGTH, "%s", range->formula.c_str());

    /* Check for valid formula. TODO. This needs stronger validation. */
    tmp_str = strchr(formula, '!');

    if (tmp_str == NULL) {
        range->ignore_cache = true;
        return;
    }
    else {
        /* Split the formulas into sheetname and row-col data. */
        *tmp_str = '\0';
        tmp_str++;
        sheetname = formula;

        /* Remove any worksheet quoting. */
        if (sheetname[0] == '\'')
            sheetname++;
        if (sheetname[strlen(sheetname) - 1] == '\'')
            sheetname[strlen(sheetname) - 1] = '\0';

        /* Check that the sheetname exists. */
        if (!get_worksheet_by_name(sheetname)) {
            LXW_WARN_FORMAT2("workbook_add_chart(): worksheet name '%s' "
                             "in chart formula '%s' doesn't exist.",
                             sheetname, range->formula.c_str());
            range->ignore_cache = true;
            return;
        }

        range->sheetname = sheetname;
        range->first_row = lxw_name_to_row(tmp_str);
        range->first_col = lxw_name_to_col(tmp_str);

        if (strchr(tmp_str, ':')) {
            /* 2D range. */
            range->last_row = lxw_name_to_row_2(tmp_str);
            range->last_col = lxw_name_to_col_2(tmp_str);
        }
        else {
            /* 1D range. */
            range->last_row = range->first_row;
            range->last_col = range->first_col;
        }

    }
}

/* Set the range dimensions and set the data cache.
 */
void workbook::_populate_range(const series_range_ptr& range)
{
    _populate_range_dimensions(range);
    _populate_range_data_cache(range);
}

/*
 * Add "cached" data to charts to provide the numCache and strCache data for
 * series and title/axis ranges.
 */
void workbook::_add_chart_cache_data()
{
    std::vector<chart*> charts;

    for (const auto& chart : ordered_charts) {
        charts.push_back(chart);
        if (chart->combined)
            charts.push_back(chart->combined.get());
    }

    for (const auto& chart : charts) {

        _populate_range(chart->title.range);
        _populate_range(chart->x_axis->title.range);
        _populate_range(chart->y_axis->title.range);

        if (chart->series_list.empty())
            continue;

        for (const auto& series : chart->series_list) {
            _populate_range(series->categories);
            _populate_range(series->values);
            _populate_range(series->title.range);
        }
    }
}

/*
 * Iterate through the worksheets and set up any chart or image drawings.
 */
void workbook::_prepare_drawings()
{
    uint16_t chart_ref_id = 0;
    uint16_t image_ref_id = 0;
    uint16_t drawing_id = 0;

    for (const auto& worksheet : worksheets) {

        if (worksheet->image_data.empty() && worksheet->chart_data.empty())
            continue;

        drawing_id++;

        for (const auto& image_options : worksheet->chart_data) {
            chart_ref_id++;
            worksheet->prepare_chart(chart_ref_id, drawing_id, image_options);
            if (image_options->chart)
                ordered_charts.push_back(image_options->chart);
        }

        for (const auto& image_options : worksheet->image_data) {

            if (image_options->image_type == LXW_IMAGE_PNG)
                has_png = true;

            if (image_options->image_type == LXW_IMAGE_JPEG)
                has_jpeg = true;

            if (image_options->image_type == LXW_IMAGE_BMP)
                has_bmp = true;

            image_ref_id++;

            worksheet->prepare_image(image_ref_id, drawing_id, image_options);
        }
    }

    drawing_count = drawing_id;
}

/*
 * Iterate through the worksheets and store any defined names used for print
 * ranges or repeat rows/columns.
 */
void workbook::_prepare_defined_names()
{
    std::string app_name; //[LXW_DEFINED_NAME_LENGTH];
    std::string range; //[LXW_DEFINED_NAME_LENGTH];
    std::string area; //[LXW_MAX_CELL_RANGE_LENGTH];
    std::string first_col;// 8
    std::string last_col; // 8

    for (const auto& worksheet : worksheets) {

        /*
         * Check for autofilter settings and store them.
         */
        if (worksheet->autofilter_.in_use) {

            app_name.append(worksheet->quoted_name);
            app_name.append("%s!_FilterDatabase");

            lxw_rowcol_to_range_abs(area,
                                    worksheet->autofilter_.first_row,
                                    worksheet->autofilter_.first_col,
                                    worksheet->autofilter_.last_row,
                                    worksheet->autofilter_.last_col);

            range.append(worksheet->quoted_name);
            range.append("!");
            range.append(area);

            /* Autofilters are the only defined name to set the hidden flag. */
            _store_defined_name("_xlnm._FilterDatabase", app_name,
                                range, worksheet->index, true);
        }

        /*
         * Check for Print Area settings and store them.
         */
        if (worksheet->print_area_.in_use) {
            app_name.clear();

            app_name.append(worksheet->quoted_name);
            app_name.append("!Print_Area");

            /* Check for print area that is the max row range. */
            if (worksheet->print_area_.first_row == 0
                && worksheet->print_area_.last_row == LXW_ROW_MAX - 1) {

                lxw_col_to_name(first_col,
                                worksheet->print_area_.first_col, false);

                lxw_col_to_name(last_col,
                                worksheet->print_area_.last_col, false);

                area = "$" + first_col + ":$" + last_col;
            }
            /* Check for print area that is the max column range. */
            else if (worksheet->print_area_.first_col == 0
                     && worksheet->print_area_.last_col == LXW_COL_MAX - 1) {

                area = "$" + std::to_string(worksheet->print_area_.first_row + 1)
                        + ":$" + std::to_string(worksheet->print_area_.last_row + 1);
            }
            else {
                lxw_rowcol_to_range_abs(area,
                                        worksheet->print_area_.first_row,
                                        worksheet->print_area_.first_col,
                                        worksheet->print_area_.last_row,
                                        worksheet->print_area_.last_col);
            }

            range =  worksheet->quoted_name + "!" + area;

            _store_defined_name("_xlnm.Print_Area", app_name,
                                range, worksheet->index, false);
        }

        /*
         * Check for repeat rows/cols. aka, Print Titles and store them.
         */
        if (worksheet->repeat_rows_.in_use || worksheet->repeat_cols_.in_use) {
            app_name.clear();
            if (worksheet->repeat_rows_.in_use
                && worksheet->repeat_cols_.in_use) {

                app_name = worksheet->quoted_name + "!Print_Titles";

                lxw_col_to_name(first_col,
                                worksheet->repeat_cols_.first_col, false);

                lxw_col_to_name(last_col,
                                worksheet->repeat_cols_.last_col, false);

                std::ostringstream ss;

                ss << worksheet->quoted_name << "!$"
                   << first_col << ":$" << last_col <<","
                   << worksheet->quoted_name << "!$"
                   << worksheet->repeat_rows_.first_row + 1 << ":$"
                   << worksheet->repeat_rows_.last_row + 1;

                range = ss.str();

                _store_defined_name("_xlnm.Print_Titles", app_name,
                                    range, worksheet->index, false);
            }
            else if (worksheet->repeat_rows_.in_use) {
                app_name = worksheet->quoted_name + "!Print_Titles";

                std::ostringstream ss;

                ss << worksheet->quoted_name << "!$"
                   << worksheet->repeat_rows_.first_row + 1 << ":$"
                   << worksheet->repeat_rows_.last_row + 1;

                range = ss.str();

                _store_defined_name("_xlnm.Print_Titles", app_name,
                                    range, worksheet->index, false);
            }
            else if (worksheet->repeat_cols_.in_use) {
                app_name = worksheet->quoted_name + "!Print_Titles";

                lxw_col_to_name(first_col,
                                worksheet->repeat_cols_.first_col, false);

                lxw_col_to_name(last_col,
                                worksheet->repeat_cols_.last_col, false);

                std::ostringstream ss;

                ss << worksheet->quoted_name << "!$"
                   << first_col << ":$"
                   << last_col;

                range = ss.str();

                _store_defined_name("_xlnm.Print_Titles", app_name,
                                    range, worksheet->index, false);
            }
        }
    }
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
void workbook::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <workbook> element.
 */
void workbook::_write_workbook()
{
    xml_attribute_list attributes = {
        {"xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main" },
        {"xmlns:r","http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    };

    lxw_xml_start_tag("workbook", attributes);
}

/*
 * Write the <fileVersion> element.
 */
void workbook::_write_file_version()
{
    xml_attribute_list attributes = {
        {"appName", "xl"},
        {"lastEdited", "4"},
        {"lowestEdited", "4"},
        {"rupBuild", "4505"}
    };
    lxw_xml_empty_tag("fileVersion", attributes);
}

/*
 * Write the <workbookPr> element.
 */
void workbook::_write_workbook_pr()
{
    xml_attribute_list attributes = {
        {"defaultThemeVersion", "124226"}
    };

    lxw_xml_empty_tag("workbookPr", attributes);
}

/*
 * Write the <workbookView> element.
 */
void workbook::_write_workbook_view()
{
    xml_attribute_list attributes = {
        {"xWindow", "240"},
        {"yWindow", "15"},
        {"windowWidth", "16095"},
        {"windowHeight", "9660"}
    };

    if (first_sheet)
        attributes.push_back({"firstSheet", std::to_string(first_sheet)});

    if (active_sheet)
        attributes.push_back({"activeTab", std::to_string(active_sheet)});

    lxw_xml_empty_tag("workbookView", attributes);
}

/*
 * Write the <bookViews> element.
 */
void workbook::_write_book_views()
{
    lxw_xml_start_tag("bookViews");

    _write_workbook_view();

    lxw_xml_end_tag("bookViews");
}

/*
 * Write the <sheet> element.
 */
void workbook::_write_sheet(const std::string& name, uint32_t sheet_id, uint8_t hidden)
{

    char r_id[LXW_MAX_ATTRIBUTE_LENGTH] = "rId1";

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", sheet_id);

    xml_attribute_list attributes = {
        {"name", name},
        {"sheetId", std::to_string(sheet_id)}
    };

    if (hidden)
        attributes.push_back({"state", "hidden"});

    attributes.push_back({"r:id", r_id});

    lxw_xml_empty_tag("sheet", attributes);
}

/*
 * Write the <sheets> element.
 */
void workbook::_write_sheets()
{
    lxw_xml_start_tag("sheets");

    for (const auto& worksheet : worksheets) {
        _write_sheet(worksheet->name, worksheet->index + 1, worksheet->hidden);
    }

    lxw_xml_end_tag("sheets");
}

/*
 * Write the <calcPr> element.
 */
void workbook::_write_calc_pr()
{
    xml_attribute_list attributes = {
        {"calcId", "124519"},
        {"fullCalcOnLoad", "1"}
    };

    lxw_xml_empty_tag("calcPr", attributes);
}

/*
 * Write the <definedName> element.
 */
void workbook::_write_defined_name(const defined_name_ptr& defined_name)
{
    xml_attribute_list attributes = {
        {"name", defined_name->name}
    };

    if (defined_name->index != -1)
        attributes.push_back({"localSheetId", std::to_string(defined_name->index)});

    if (defined_name->hidden)
        attributes.push_back({"hidden", "1"});

    lxw_xml_data_element("definedName", defined_name->formula, attributes);
}

/*
 * Write the <definedNames> element.
 */
void workbook::_write_defined_names()
{
    if (defined_names.empty())
        return;

    lxw_xml_start_tag("definedNames");

    for (const auto& defined_name : defined_names) {
        _write_defined_name(defined_name);
    }

    lxw_xml_end_tag("definedNames");
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Assemble and write the XML file.
 */
void workbook::assemble_xml_file()
{
    /* Prepare workbook and sub-objects for writing. */
    _prepare_workbook();

    /* Write the XML declaration. */
    _xml_declaration();

    /* Write the root workbook element. */
    _write_workbook();

    /* Write the XLSX file version. */
    _write_file_version();

    /* Write the workbook properties. */
    _write_workbook_pr();

    /* Write the workbook view properties. */
    _write_book_views();

    /* Write the worksheet names and ids. */
    _write_sheets();

    /* Write the workbook defined names. */
    _write_defined_names();

    /* Write the workbook calculation properties. */
    _write_calc_pr();

    /* Close the workbook tag. */
    lxw_xml_end_tag("workbook");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Create a new workbook object.
 */

workbook::workbook(const std::string& file, const workbook_options& options) : filename(file)
{
    workbook_new_opt(options);
}

workbook::~workbook()
{
    for(auto p : used_xf_formats.order_list)
    {
        delete p.first;
        delete p.second;
    }
}

/*
 * Create a new workbook object with options.
 */
void workbook::workbook_new_opt(const workbook_options& options)
{
    /* Add the shared strings table. */
    sst = std::make_shared<xlsxwriter::sst>();

    /* Add the default cell format. */
    auto format = add_format();

    /* Initialize its index. */
    format->get_xf_index();

    this->options.constant_memory = options.constant_memory;
    this->options.tmpdir = options.tmpdir;
}

/*
 * Add a new worksheet to the Excel workbook.
 */
worksheet* workbook::add_worksheet(const std::string& sheetname)
{
    lxw_worksheet_init_data init_data = {};
    std::string new_name;

    if (!sheetname.empty()) {
        /* Use the user supplied name. */
        if (sheetname.size() > LXW_SHEETNAME_MAX) {
            return nullptr;
        }
        else {
            init_data.name = sheetname;
            init_data.quoted_name = lxw_quote_sheetname(sheetname);
        }
    }
    else {
        /* Use the default SheetN name. */
        new_name = "Sheet" + std::to_string(num_sheets + 1);
        init_data.name = new_name;
        init_data.quoted_name = new_name;
    }

    /* Check if the worksheet name is already in use. */
    if (get_worksheet_by_name(init_data.name)) {
        LXW_WARN_FORMAT1("workbook_add_worksheet(): worksheet name '%s' "
                         "already exists.", init_data.name.c_str());
        return nullptr;
    }

    /* Initialize the metadata to pass to the worksheet. */
    init_data.hidden = 0;
    init_data.index = num_sheets;
    init_data.sst = sst;
    init_data.optimize = options.constant_memory;
    init_data.active_sheet = &active_sheet;
    init_data.first_sheet = &first_sheet;
    init_data.tmpdir = options.tmpdir;

    /* Create a new worksheet object. */
    worksheet_ptr worksheet = std::make_shared<xlsxwriter::worksheet>(&init_data);

    num_sheets++;
    worksheets.push_back(worksheet);

    /* Store the worksheet so we can look it up by name. */    
    worksheet_names.insert(std::make_pair(init_data.name, worksheet));

    return worksheet.get();
}

/*
 * Add a new chart to the Excel workbook.
 */
chart* workbook::add_chart(uint8_t type)
{
    /* Create a new chart object. */
    chart_ptr chart;
    switch(type) {
    case LXW_CHART_AREA:
    case LXW_CHART_AREA_STACKED:
    case LXW_CHART_AREA_STACKED_PERCENT:
        chart = std::dynamic_pointer_cast<xlsxwriter::chart>(std::make_shared<chart_area>(type));
        break;
    case LXW_CHART_BAR:
    case LXW_CHART_BAR_STACKED:
    case LXW_CHART_BAR_STACKED_PERCENT:
        chart = std::dynamic_pointer_cast<xlsxwriter::chart>(std::make_shared<chart_bar>(type));
        break;
    case LXW_CHART_COLUMN:
    case LXW_CHART_COLUMN_STACKED:
    case LXW_CHART_COLUMN_STACKED_PERCENT:
        chart = std::dynamic_pointer_cast<xlsxwriter::chart>(std::make_shared<chart_column>(type));
        break;
    case LXW_CHART_DOUGHNUT:
        chart = std::dynamic_pointer_cast<xlsxwriter::chart>(std::make_shared<chart_doughtnut>(type));
        break;
    case LXW_CHART_LINE:
        chart = std::dynamic_pointer_cast<xlsxwriter::chart>(std::make_shared<chart_line>(type));
        break;
    case LXW_CHART_PIE:
        chart = std::dynamic_pointer_cast<xlsxwriter::chart>(std::make_shared<chart_pie>(type));
        break;
    case LXW_CHART_SCATTER:
    case LXW_CHART_SCATTER_STRAIGHT:
    case LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS:
    case LXW_CHART_SCATTER_SMOOTH:
    case LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS:
        chart = std::dynamic_pointer_cast<xlsxwriter::chart>(std::make_shared<chart_scatter>(type));
        break;
    case LXW_CHART_RADAR:
    case LXW_CHART_RADAR_WITH_MARKERS:
    case LXW_CHART_RADAR_FILLED:
        chart = std::dynamic_pointer_cast<xlsxwriter::chart>(std::make_shared<chart_radar>(type));
        break;
    default:
        return nullptr;
    }
    charts.push_back(chart);
    return chart.get();
}

/*
 * Add a new format to the Excel workbook.
 */
format* workbook::add_format()
{
    /* Create a new format object. */
    format_ptr format = new xlsxwriter::format();

    format->xf_format_indices = &used_xf_formats;
    format->num_xf_formats = &num_xf_formats;

    formats.push_back(format);

    return format;
}

/*
 * Call finalization code and close file.
 */
lxw_error workbook::close()
{
    lxw_error error = LXW_NO_ERROR;

    /* Add a default worksheet if non have been added. */
    if (worksheets.empty())
        add_worksheet();

    /* Ensure that at least one worksheet has been selected. */
    if (active_sheet == 0) {
        const auto& sheet = worksheets.front();
        sheet->selected = 1;
        sheet->hidden = 0;
    }

    /* Set the active sheet. */
    for (const auto& sheet : worksheets) {
        if (sheet->index == active_sheet)
            sheet->active = 1;
    }

    /* Set the defined names for the worksheets such as Print Titles. */
    _prepare_defined_names();

    /* Prepare the drawings, charts and images. */
    _prepare_drawings();

    /* Add cached data to charts. */
    _add_chart_cache_data();

    /* Create a packager object to assemble sub-elements into a zip file. */
    std::shared_ptr<packager> pkger = std::make_shared<packager>(filename, options.tmpdir);

    /* Set the workbook object in the packager. */
    pkger->workbook = this;

    /* Assemble all the sub-files in the xlsx package. */
    error = (lxw_error)pkger->create_package();

    /* Error and non-error conditions fall through to the cleanup code. */
    if (error == LXW_ERROR_CREATING_TMPFILE) {
        std::cerr << "[ERROR] workbook_close(): "
             << "Error creating tmpfile(s) to assemble '" << filename << "'. "
             << "Error = " << strerror(errno) << std::endl;
    }

    /* If LXW_ERROR_ZIP_FILE_OPERATION then errno is set by zlib. */
    if (error == LXW_ERROR_ZIP_FILE_OPERATION) {
        std::cerr << "[ERROR] workbook_close(): "
             << "Zlib error while creating xlsx file '" << filename << "'. "
             << "Error = " << strerror(errno) << std::endl;
    }

    /* The next 2 error conditions don't set errno. */
    if (error == LXW_ERROR_ZIP_FILE_ADD) {
        std::cerr << "[ERROR] workbook_close(): "
            << "Zlib error adding file to xlsx file '"<< filename <<"'." << std::endl;
    }

    if (error == LXW_ERROR_ZIP_CLOSE) {
        std::cerr << "[ERROR] workbook_close(): "
                  << "Zlib error closing xlsx file ' " << filename <<"'." << std::endl;
    }

    return error;
}

/*
 * Create a defined name in Excel. We handle global/workbook level names and
 * local/worksheet names.
 */
lxw_error workbook::define_name(const std::string& name, const std::string& formula)
{
    return _store_defined_name(name, "", formula, -1, false);
}

/*
 * Set the document properties such as Title, Author etc.
 */
lxw_error workbook::set_properties(const doc_properties& user_props)
{
    properties = user_props;

    return LXW_NO_ERROR;
}

/*
 * Set a string custom document property.
 */
lxw_error workbook::set_custom_property_string(const std::string& name, const std::string& value)
{
    if (name.empty()) {
        LXW_WARN_FORMAT("workbook_set_custom_property_string(): "
                        "parameter 'name' cannot be NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (value.empty()) {
        LXW_WARN_FORMAT("workbook_set_custom_property_string(): "
                        "parameter 'value' cannot be NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (name.size() > 255) {
        LXW_WARN_FORMAT("workbook_set_custom_property_string(): parameter "
                        "'name' exceeds Excel length limit of 255.");
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    if (value.size() > 255) {
        LXW_WARN_FORMAT("workbook_set_custom_property_string(): parameter "
                        "'value' exceeds Excel length limit of 255.");
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Create a struct to hold the custom property. */
    custom_property_ptr custom_property = std::make_shared<lxw_custom_property>();

    custom_property->name = name;
    custom_property->u.string = value;
    custom_property->type = LXW_CUSTOM_STRING;

    custom_properties.push_back(custom_property);

    return LXW_NO_ERROR;
}

/*
 * Set a double number custom document property.
 */
lxw_error workbook::set_custom_property_number(const std::string& name, double value)
{
    if (name.empty()) {
        LXW_WARN_FORMAT("workbook_set_custom_property_number(): parameter "
                        "'name' cannot be NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (name.size() > 255) {
        LXW_WARN_FORMAT("workbook_set_custom_property_number(): parameter "
                        "'name' exceeds Excel length limit of 255.");
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Create a struct to hold the custom property. */
    custom_property_ptr custom_property = std::make_shared<lxw_custom_property>();

    custom_property->name = name;
    custom_property->u.number = value;
    custom_property->type = LXW_CUSTOM_DOUBLE;

    custom_properties.push_back(custom_property);

    return LXW_NO_ERROR;
}

/*
 * Set a integer number custom document property.
 */
lxw_error workbook::set_custom_property_integer(const std::string& name, int32_t value)
{
    if (name.empty()) {
        LXW_WARN_FORMAT("workbook_set_custom_property_integer(): parameter "
                        "'name' cannot be NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (name.size() > 255) {
        LXW_WARN_FORMAT("workbook_set_custom_property_integer(): parameter "
                        "'name' exceeds Excel length limit of 255.");
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Create a struct to hold the custom property. */
    custom_property_ptr custom_property = std::make_shared<lxw_custom_property>();

    custom_property->name = name;
    custom_property->u.integer = value;
    custom_property->type = LXW_CUSTOM_INTEGER;

    custom_properties.push_back(custom_property);

    return LXW_NO_ERROR;
}

/*
 * Set a boolean custom document property.
 */
lxw_error workbook::set_custom_property_boolean(const std::string& name, bool value)
{
    if (name.empty()) {
        LXW_WARN_FORMAT("workbook_set_custom_property_boolean(): parameter "
                        "'name' cannot be NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (name.size() > 255) {
        LXW_WARN_FORMAT("workbook_set_custom_property_boolean(): parameter "
                        "'name' exceeds Excel length limit of 255.");
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;
    }

    /* Create a struct to hold the custom property. */
    custom_property_ptr custom_property = std::make_shared<lxw_custom_property>();

    custom_property->name = name;
    custom_property->u.boolean = value;
    custom_property->type = LXW_CUSTOM_BOOLEAN;

    custom_properties.push_back(custom_property);

    return LXW_NO_ERROR;
}

/*
 * Set a datetime custom document property.
 */
lxw_error workbook::set_custom_property_datetime(const std::string& name, lxw_datetime *datetime)
{
    if (name.empty()) {
        LXW_WARN_FORMAT("workbook_set_custom_property_datetime(): parameter "
                        "'name' cannot be NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (name.size() > 255) {
        LXW_WARN_FORMAT("workbook_set_custom_property_datetime(): parameter "
                        "'name' exceeds Excel length limit of 255.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    if (!datetime) {
        LXW_WARN_FORMAT("workbook_set_custom_property_datetime(): parameter "
                        "'datetime' cannot be NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Create a struct to hold the custom property. */
    custom_property_ptr custom_property = std::make_shared<lxw_custom_property>();

    custom_property->name = name;

    memcpy(&(custom_property->u.datetime), datetime, sizeof(lxw_datetime));
    custom_property->type = LXW_CUSTOM_DATETIME;

    custom_properties.push_back(custom_property);

    return LXW_NO_ERROR;
}

worksheet* workbook::get_worksheet_by_name(const std::string& name)
{
    if (name.empty())
        return nullptr;

    auto it = worksheet_names.find(name);
    if (it != worksheet_names.end())
        return it->second.get();
    else
        return nullptr;
}

} // namespace xlsxwriter
