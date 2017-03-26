/*****************************************************************************
 * worksheet - A library for creating Excel XLSX worksheet files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <ctype.h>

#include "xmlwriter.hpp"
#include "worksheet.hpp"
#include "format.hpp"
#include "utility.hpp"
#include "relationships.hpp"
#include <iostream>
#include <iomanip>

#define LXW_STR_MAX      32767
#define LXW_BUFFER_SIZE  4096
#define LXW_PORTRAIT     1
#define LXW_LANDSCAPE    0
#define LXW_PRINT_ACROSS 1

namespace xlsxwriter {

/*
 * Forward declarations.
 */
void _write_rows();
int _row_cmp(lxw_row *row1, lxw_row *row2);
int _cell_cmp(lxw_cell *cell1, lxw_cell *cell2);

LXW_RB_GENERATE_ROW(lxw_table_rows, lxw_row, tree_pointers, _row_cmp);
LXW_RB_GENERATE_CELL(lxw_table_cells, lxw_cell, tree_pointers, _cell_cmp);

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Find but don't create a row object for a given row number.
 */
lxw_row * worksheet::find_row(lxw_row_t row_num)
{
    lxw_row row;

    row.row_num = row_num;

    return RB_FIND(lxw_table_rows, table, &row);
}

/*
 * Find but don't create a cell object for a given row object and col number.
 */
lxw_cell * worksheet::find_cell(lxw_row *row, lxw_col_t col_num)
{
    lxw_cell cell;

    if (!row)
        return NULL;

    cell.col_num = col_num;

    return RB_FIND(lxw_table_cells, row->cells, &cell);
}

/*
 * Create a new worksheet object.
 */
worksheet::worksheet(lxw_worksheet_init_data *init_data)
{
    table = new lxw_table_rows();
    RB_INIT(table);

    hyperlinks = new lxw_table_rows();
    RB_INIT(hyperlinks);

    /* Initialize the cached rows. */
    table->cached_row_num = LXW_ROW_MAX + 1;
    hyperlinks->cached_row_num = LXW_ROW_MAX + 1;

    if (init_data && init_data->optimize) {
        array = new lxw_cell *[LXW_COL_MAX]();
    }

    col_options = new lxw_col_options*[LXW_COL_META_MAX]();
    col_options_max = LXW_COL_META_MAX;

    col_formats = new xlsxwriter::format*[LXW_COL_META_MAX]();
    col_formats_max = LXW_COL_META_MAX;

    optimize_row = new lxw_row();
    optimize_row->height = LXW_DEF_ROW_HEIGHT;

    if (init_data && init_data->optimize) {
        FILE *tmpfile;

        if (init_data)
            tmpfile = lxw_tmpfile(init_data->tmpdir.c_str());
        else
            tmpfile = lxw_tmpfile(NULL);

        if (!tmpfile) {
            throw std::string("Error creating tmpfile() for worksheet in "
                      "'constant_memory' mode.");
        }

        optimize_tmpfile = tmpfile;
        file = optimize_tmpfile;
    }

    /* Initialize the worksheet dimensions. */
    dim_rowmax = 0;
    dim_colmax = 0;
    dim_rowmin = LXW_ROW_MAX;
    dim_colmin = LXW_COL_MAX;

    default_row_height = LXW_DEF_ROW_HEIGHT;
    default_row_pixels = 20;
    default_col_pixels = 64;

    /* Initialize the page setup properties. */
    fit_height = 0;
    fit_width = 0;
    page_start = 0;
    print_scale = 100;
    fit_page = 0;
    orientation = true;
    page_order = 0;
    page_setup_changed = false;
    page_view = false;
    paper_size = 0;
    vertical_dpi = 0;
    horizontal_dpi = 0;
    margin_left = 0.7;
    margin_right = 0.7;
    margin_top = 0.75;
    margin_bottom = 0.75;
    margin_header = 0.3;
    margin_footer = 0.3;
    print_gridlines = 0;
    screen_gridlines = 1;
    print_options_changed = 0;
    zoom = 100;
    zoom_scale_normal = true;
    show_zeros = true;
    outline_on = true;
    tab_color = LXW_COLOR_UNSET;

    if (init_data) {
        name = init_data->name;
        quoted_name = init_data->quoted_name;
        tmpdir = init_data->tmpdir;
        index = init_data->index;
        hidden = init_data->hidden;
        sst = init_data->sst;
        optimize = init_data->optimize;
        active_sheet = init_data->active_sheet;
        first_sheet = init_data->first_sheet;
    }
}

/*
 * Free a worksheet cell.
 */
void _free_cell(lxw_cell *cell)
{
    if (!cell)
        return;

    if (cell->type != NUMBER_CELL && cell->type != STRING_CELL
        && cell->type != BLANK_CELL && cell->type != BOOLEAN_CELL) {

        delete cell->u.string;
    }

    delete cell->user_data1;
    delete cell->user_data2;

    delete cell;
}

/*
 * Free a worksheet row.
 */
void
_free_row(lxw_row *row)
{
    lxw_cell *cell;
    lxw_cell *next_cell;

    if (!row)
        return;

    for (cell = RB_MIN(lxw_table_cells, row->cells); cell; cell = next_cell) {
        next_cell = RB_NEXT(lxw_table_cells, row->cells, cell);
        RB_REMOVE(lxw_table_cells, row->cells, cell);
        _free_cell(cell);
    }

    free(row->cells);
    free(row);
}

/*
 * Create a new worksheet row object.
 */
lxw_row *
_new_row(lxw_row_t row_num)
{
    lxw_row *row = new lxw_row();

    if (row) {
        row->row_num = row_num;
        row->cells = new lxw_table_cells();
        row->height = LXW_DEF_ROW_HEIGHT;

        if (row->cells)
            RB_INIT(row->cells);
        else
            LXW_MEM_ERROR();
    }
    else {
        LXW_MEM_ERROR();
    }

    return row;
}

/*
 * Create a new worksheet number cell object.
 */
lxw_cell * _new_number_cell(lxw_row_t row_num, lxw_col_t col_num, double value, xlsxwriter::format *format)
{
    lxw_cell *cell = new lxw_cell();

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = NUMBER_CELL;
    cell->format = format;
    cell->u.number = value;

    return cell;
}

/*
 * Create a new worksheet string cell object.
 */
lxw_cell * _new_string_cell(lxw_row_t row_num,
                 lxw_col_t col_num, int32_t string_id, std::string *sst_string,
                 xlsxwriter::format *format)
{
    lxw_cell *cell = new lxw_cell();

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = STRING_CELL;
    cell->format = format;
    cell->u.string_id = string_id;
    cell->sst_string = sst_string;

    return cell;
}

/*
 * Create a new worksheet inline_string cell object.
 */
lxw_cell *
_new_inline_string_cell(lxw_row_t row_num,
                        lxw_col_t col_num, std::string *string, xlsxwriter::format *format)
{
    lxw_cell *cell = new lxw_cell();

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = INLINE_STRING_CELL;
    cell->format = format;
    cell->u.string = string;

    return cell;
}

/*
 * Create a new worksheet formula cell object.
 */
lxw_cell *
_new_formula_cell(lxw_row_t row_num,
                  lxw_col_t col_num, std::string *formula, xlsxwriter::format *format)
{
    lxw_cell *cell = new lxw_cell();

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = FORMULA_CELL;
    cell->format = format;
    cell->u.string = formula;

    return cell;
}

/*
 * Create a new worksheet array formula cell object.
 */
lxw_cell *
_new_array_formula_cell(lxw_row_t row_num, lxw_col_t col_num, std::string *formula,
                        std::string *range, xlsxwriter::format *format)
{
    lxw_cell *cell = new lxw_cell();

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = ARRAY_FORMULA_CELL;
    cell->format = format;
    cell->u.string = formula;
    cell->user_data1 = range;

    return cell;
}

/*
 * Create a new worksheet blank cell object.
 */
lxw_cell *
_new_blank_cell(lxw_row_t row_num, lxw_col_t col_num, xlsxwriter::format *format)
{
    lxw_cell *cell = new lxw_cell();

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = BLANK_CELL;
    cell->format = format;

    return cell;
}

/*
 * Create a new worksheet boolean cell object.
 */
lxw_cell *
_new_boolean_cell(lxw_row_t row_num, lxw_col_t col_num, int value,
                  xlsxwriter::format *format)
{
    lxw_cell *cell = new lxw_cell();

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = BOOLEAN_CELL;
    cell->format = format;
    cell->u.number = value;

    return cell;
}

/*
 * Create a new worksheet hyperlink cell object.
 */
lxw_cell *
_new_hyperlink_cell(lxw_row_t row_num, lxw_col_t col_num,
                    enum cell_types link_type, std::string *url, std::string *string,
                    std::string *tooltip)
{
    lxw_cell *cell = new lxw_cell();

    cell->row_num = row_num;
    cell->col_num = col_num;
    cell->type = link_type;
    cell->u.string = url;
    cell->user_data1 = string;
    cell->user_data2 = tooltip;

    return cell;
}

/*
 * Get or create the row object for a given row number.
 */
lxw_row *
_get_row_list(struct lxw_table_rows *table, lxw_row_t row_num)
{
    lxw_row *row;
    lxw_row *existing_row;

    if (table->cached_row_num == row_num)
        return table->cached_row;

    /* Create a new row and try and insert it. */
    row = _new_row(row_num);
    existing_row = RB_INSERT(lxw_table_rows, table, row);

    /* If existing_row is not NULL, then it already existed. Free new row */
    /* and return existing_row. */
    if (existing_row) {
        _free_row(row);
        row = existing_row;
    }

    table->cached_row = row;
    table->cached_row_num = row_num;

    return row;
}

/*
 * Get or create the row object for a given row number.
 */
lxw_row * worksheet::_get_row(lxw_row_t row_num)
{
    lxw_row *row;

    if (!optimize) {
        row = _get_row_list(table, row_num);
        return row;
    }
    else {
        if (row_num < optimize_row->row_num) {
            return NULL;
        }
        else if (row_num == optimize_row->row_num) {
            return optimize_row;
        }
        else {
            /* Flush row. */
            write_single_row();
            row = optimize_row;
            row->row_num = row_num;
            return row;
        }
    }
}

/*
 * Insert a cell object in the cell list of a row object.
 */
void _insert_cell_list(lxw_table_cells *cell_list,
                  lxw_cell *cell, lxw_col_t col_num)
{
    lxw_cell *existing_cell;

    cell->col_num = col_num;

    existing_cell = RB_INSERT(lxw_table_cells, cell_list, cell);

    /* If existing_cell is not NULL, then that cell already existed. */
    /* Remove existing_cell and add new one in again. */
    if (existing_cell) {
        RB_REMOVE(lxw_table_cells, cell_list, existing_cell);

        /* Add it in again. */
        RB_INSERT(lxw_table_cells, cell_list, cell);
        _free_cell(existing_cell);
    }

    return;
}

/*
 * Insert a cell object into the cell list or array.
 */
void worksheet::_insert_cell(lxw_row_t row_num, lxw_col_t col_num, lxw_cell *cell)
{
    lxw_row *row = _get_row(row_num);

    if (!optimize) {
        row->data_changed = true;
        _insert_cell_list(row->cells, cell, col_num);
    }
    else {
        if (row) {
            row->data_changed = true;

            /* Overwrite an existing cell if necessary. */
            if (array[col_num])
                _free_cell(array[col_num]);

            array[col_num] = cell;
        }
    }
}

/*
 * Insert a hyperlink object into the hyperlink list.
 */
void worksheet::_insert_hyperlink(lxw_row_t row_num, lxw_col_t col_num, lxw_cell *link)
{
    lxw_row *row = _get_row_list(hyperlinks, row_num);

    _insert_cell_list(row->cells, link, col_num);
}

/*
 * Next power of two for column reallocs. Taken from bithacks in the public
 * domain.
 */
lxw_col_t
_next_power_of_two(uint16_t col)
{
    col--;
    col |= col >> 1;
    col |= col >> 2;
    col |= col >> 4;
    col |= col >> 8;
    col++;

    return col;
}

/*
 * Check that row and col are within the allowed Excel range and store max
 * and min values for use in other methods/elements.
 *
 * The ignore_row/ignore_col flags are used to indicate that we wish to
 * perform the dimension check without storing the value.
 */
lxw_error worksheet::_check_dimensions(lxw_row_t row_num, lxw_col_t col_num, int8_t ignore_row, int8_t ignore_col)
{
    if (row_num >= LXW_ROW_MAX)
        return LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;

    if (col_num >= LXW_COL_MAX)
        return LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;

    /* In optimization mode we don't change dimensions for rows that are */
    /* already written. */
    if (!ignore_row && !ignore_col && optimize) {
        if (row_num < optimize_row->row_num)
            return LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE;
    }

    if (!ignore_row) {
        if (row_num < dim_rowmin)
            dim_rowmin = row_num;
        if (row_num > dim_rowmax)
            dim_rowmax = row_num;
    }

    if (!ignore_col) {
        if (col_num < dim_colmin)
            dim_colmin = col_num;
        if (col_num > dim_colmax)
            dim_colmax = col_num;
    }

    return LXW_NO_ERROR;
}

/*
 * Comparator for the row structure red/black tree.
 */
int
_row_cmp(lxw_row *row1, lxw_row *row2)
{
    if (row1->row_num > row2->row_num)
        return 1;
    if (row1->row_num < row2->row_num)
        return -1;
    return 0;
}

/*
 * Comparator for the cell structure red/black tree.
 */
int
_cell_cmp(lxw_cell *cell1, lxw_cell *cell2)
{
    if (cell1->col_num > cell2->col_num)
        return 1;
    if (cell1->col_num < cell2->col_num)
        return -1;
    return 0;
}

/*
 * Hash a worksheet password. Based on the algorithm provided by Daniel Rentz
 * of OpenOffice.
 */
uint16_t
_hash_password(const char *password)
{
    size_t count;
    uint8_t i;
    uint16_t hash = 0x0000;

    count = strlen(password);

    for (i = 0; i < count; i++) {
        uint32_t low_15;
        uint32_t high_15;
        uint32_t letter = password[i] << (i + 1);

        low_15 = letter & 0x7fff;
        high_15 = letter & (0x7fff << 15);
        high_15 = high_15 >> 15;
        letter = low_15 | high_15;

        hash ^= letter;
    }

    hash ^= count;
    hash ^= 0xCE4B;

    return hash;
}

/*
 * Simple replacement for libgen.h basename() for compatibility with MSVC. It
 * handles forward and back slashes. It doesn't copy exactly the return
 * format of basename().
 */
std::string lxw_basename(const std::string& path)
{
    if (path.empty())
        return std::string();

    const char* forward_slash = strrchr(path.c_str(), '/');
    const char* back_slash = strrchr(path.c_str(), '\\');

    if (!forward_slash && !back_slash)
        return path;

    if (forward_slash > back_slash)
        return path.substr( forward_slash + 1 - path.c_str());
    else
        return path.substr( back_slash + 1 - path.c_str());
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/
/*
 * Write the XML declaration.
 */
void worksheet::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <worksheet> element.
 */
void worksheet::_write_worksheet()
{
    xml_attribute_list attributes = {
        {"xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
        {"xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    };
    lxw_xml_start_tag("worksheet", attributes);
}

/*
 * Write the <dimension> element.
 */
void worksheet::_write_dimension()
{
    std::string ref;
    lxw_row_t dim_rowmin = dim_rowmin;
    lxw_row_t dim_rowmax = dim_rowmax;
    lxw_col_t dim_colmin = dim_colmin;
    lxw_col_t dim_colmax = dim_colmax;

    if (dim_rowmin == LXW_ROW_MAX && dim_colmin == LXW_COL_MAX) {
        /* If the rows and cols are still the defaults then no dimensions have
         * been set and we use the default range "A1". */
        lxw_rowcol_to_range(ref, 0, 0, 0, 0);
    }
    else if (dim_rowmin == LXW_ROW_MAX && dim_colmin != LXW_COL_MAX) {
        /* If the rows aren't set but the columns are then the dimensions have
         * been changed via set_column(). */
        lxw_rowcol_to_range(ref, 0, dim_colmin, 0, dim_colmax);
    }
    else {
        lxw_rowcol_to_range(ref, dim_rowmin, dim_colmin, dim_rowmax,
                            dim_colmax);
    }

    xml_attribute_list attributes = {
        {"ref", ref}
    };

    lxw_xml_empty_tag("dimension", attributes);
}

/*
 * Write the <pane> element for freeze panes.
 */
void worksheet::_write_freeze_panes()
{
    xml_attribute_list attributes;

    std::shared_ptr<lxw_selection> selection;
    std::shared_ptr<lxw_selection> user_selection;
    lxw_row_t row = panes.first_row;
    lxw_col_t col = panes.first_col;
    lxw_row_t top_row = panes.top_row;
    lxw_col_t left_col = panes.left_col;

    std::string row_cell;
    std::string col_cell;
    std::string top_left_cell;
    std::string active_pane;

    /* If there is a user selection we remove it from the list and use it. */
    if (!selections.empty()) {
        user_selection = selections.front();
        selections.erase(selections.begin());
    }
    else {
        /* or else create a new blank selection. */
        user_selection = std::make_shared<lxw_selection>();
    }

    lxw_rowcol_to_cell(top_left_cell, top_row, left_col);

    /* Set the active pane. */
    if (row && col) {
        active_pane = "bottomRight";

        lxw_rowcol_to_cell(row_cell, row, 0);
        lxw_rowcol_to_cell(col_cell, 0, col);

        selection = std::make_shared<lxw_selection>();
        if (selection) {
            selection->pane = "topRight";
            selection->active_cell = col_cell;
            selection->sqref = col_cell;

            selections.push_back(selection);
        }

        selection = std::make_shared<lxw_selection>();
        if (selection) {
            selection->pane = "bottomLeft";
            selection->active_cell = row_cell;
            selection->sqref = row_cell;

            selections.push_back(selection);
        }

        selection = std::make_shared<lxw_selection>();
        if (selection) {
            selection->pane = "bottomRight";
            selection->active_cell, user_selection->active_cell;
            selection->sqref = user_selection->sqref;

            selections.push_back(selection);
        }
    }
    else if (col) {
        active_pane = "topRight";

        selection = std::make_shared<lxw_selection>();
        if (selection) {
            selection->pane = "topRight";
            selection->active_cell = user_selection->active_cell;
            selection->sqref = user_selection->sqref;

            selections.push_back(selection);
        }
    }
    else {
        active_pane = "bottomLeft";

        selection = std::make_shared<lxw_selection>();
        if (selection) {
            selection->pane = "bottomLeft";
            selection->active_cell = user_selection->active_cell;
            selection->sqref = user_selection->sqref;

            selections.push_back(selection);
        }
    }

    if (col)
        attributes.push_back({"xSplit", std::to_string(col)});

    if (row)
        attributes.push_back({"ySplit", std::to_string(row)});

    attributes.push_back({"topLeftCell", top_left_cell});
    attributes.push_back({"activePane", active_pane});

    if (panes.type == FREEZE_PANES)
        attributes.push_back({"state", "frozen"});
    else if (panes.type == FREEZE_SPLIT_PANES)
        attributes.push_back({"state", "frozenSplit"});

    lxw_xml_empty_tag("pane", attributes);
}

/*
 * Convert column width from user units to pane split width.
 */
uint32_t worksheet::_calculate_x_split_width(double x_split) const
{
    uint32_t width;
    uint32_t pixels;
    uint32_t points;
    uint32_t twips;
    double max_digit_width = 7.0;       /* For Calabri 11. */
    double padding = 5.0;

    /* Convert to pixels. */
    if (x_split < 1.0) {
        pixels = (uint32_t) (x_split * (max_digit_width + padding) + 0.5);
    }
    else {
        pixels = (uint32_t) (x_split * max_digit_width + 0.5) + 5;
    }

    /* Convert to points. */
    points = (pixels * 3) / 4;

    /* Convert to twips (twentieths of a point). */
    twips = points * 20;

    /* Add offset/padding. */
    width = twips + 390;

    return width;
}

/*
 * Write the <pane> element for split panes.
 */
void worksheet::_write_split_panes()
{
    xml_attribute_list attributes;

    std::shared_ptr<lxw_selection> selection;
    std::shared_ptr<lxw_selection> user_selection;
    lxw_row_t row = panes.first_row;
    lxw_col_t col = panes.first_col;
    lxw_row_t top_row = panes.top_row;
    lxw_col_t left_col = panes.left_col;
    double x_split = panes.x_split;
    double y_split = panes.y_split;
    uint8_t has_selection = false;

    std::string row_cell;//[LXW_MAX_CELL_NAME_LENGTH];
    std::string col_cell;//[LXW_MAX_CELL_NAME_LENGTH];
    std::string top_left_cell;//[LXW_MAX_CELL_NAME_LENGTH];
    std::string active_pane;//[LXW_PANE_NAME_LENGTH];

    /* If there is a user selection we remove it from the list and use it. */
    if (!selections.empty()) {
        user_selection = selections.front();
        selections.erase(selections.begin());
        has_selection = true;
    }
    else {
        /* or else create a new blank selection. */
        user_selection = std::make_shared<lxw_selection>();
    }


    /* Convert the row and col to 1/20 twip units with padding. */
    if (y_split > 0.0)
        y_split = (uint32_t) (20 * y_split + 300);

    if (x_split > 0.0)
        x_split = _calculate_x_split_width(x_split);

    /* For non-explicit topLeft definitions, estimate the cell offset based on
     * the pixels dimensions. This is only a workaround and doesn't take
     * adjusted cell dimensions into account.
     */
    if (top_row == row && left_col == col) {
        top_row = (lxw_row_t) (0.5 + (y_split - 300.0) / 20.0 / 15.0);
        left_col = (lxw_col_t) (0.5 + (x_split - 390.0) / 20.0 / 3.0 / 16.0);
    }

    lxw_rowcol_to_cell(top_left_cell, top_row, left_col);

    /* If there is no selection set the active cell to the top left cell. */
    if (!has_selection) {
        user_selection->active_cell =  top_left_cell;
        user_selection->sqref = top_left_cell;
    }

    /* Set the active pane. */
    if (y_split > 0.0 && x_split > 0.0) {
        active_pane = "bottomRight";

        lxw_rowcol_to_cell(row_cell, top_row, 0);
        lxw_rowcol_to_cell(col_cell, 0, left_col);

        selection = std::make_shared<lxw_selection>();
        if (selection) {
            selection->pane = "topRight";
            selection->active_cell = col_cell;
            selection->sqref = col_cell;

            selections.push_back(selection);
        }

        selection = std::make_shared<lxw_selection>();
        if (selection) {
            selection->pane = "bottomLeft";
            selection->active_cell = row_cell;
            selection->sqref = row_cell;

            selections.push_back(selection);
        }

        selection = std::make_shared<lxw_selection>();
        if (selection) {
            selection->pane = "bottomRight";
            selection->active_cell = user_selection->active_cell;
            selection->sqref = user_selection->sqref;

            selections.push_back(selection);
        }
    }
    else if (x_split > 0.0) {
        active_pane = "topRight";

        selection = std::make_shared<lxw_selection>();
        if (selection) {
            selection->pane = "topRight";
            selection->active_cell = user_selection->active_cell;
            selection->sqref = user_selection->sqref;

            selections.push_back(selection);
        }
    }
    else {
        active_pane = "bottomLeft";

        selection = std::make_shared<lxw_selection>();
        if (selection) {
            selection->pane = "bottomLeft";
            selection->active_cell = user_selection->active_cell;
            selection->sqref = user_selection->sqref;

            selections.push_back(selection);
        }
    }

    if (x_split > 0.0)
        attributes.push_back({"xSplit", std::to_string(x_split)});

    if (y_split > 0.0)
        attributes.push_back({"ySplit", std::to_string(y_split)});

    attributes.push_back({"topLeftCell", top_left_cell});

    if (has_selection)
        attributes.push_back({"activePane", active_pane});

    lxw_xml_empty_tag("pane", attributes);
}

/*
 * Write the <selection> element.
 */
void worksheet::_write_selection(const std::shared_ptr<lxw_selection>& selection)
{
    xml_attribute_list attributes;

    if (!selection->pane.empty())
        attributes.push_back({"pane", selection->pane});

    if (!selection->active_cell.empty())
        attributes.push_back({"activeCell", selection->active_cell});

    if (!selection->sqref.empty())
        attributes.push_back({"sqref", selection->sqref});

    lxw_xml_empty_tag("selection", attributes);
}

/*
 * Write the <selection> elements.
 */
void worksheet::_write_selections()
{
    for (const auto& selection : selections) {
       _write_selection(selection);
    }
}

/*
 * Write the frozen or split <pane> elements.
 */
void worksheet::_write_panes()
{
    if (panes.type == NO_PANES)
        return;

    else if (panes.type == FREEZE_PANES)
       _write_freeze_panes();

    else if (panes.type == FREEZE_SPLIT_PANES)
       _write_freeze_panes();

    else if (panes.type == SPLIT_PANES)
       _write_split_panes();
}

/*
 * Write the <sheetView> element.
 */
void worksheet::_write_sheet_view()
{
    xml_attribute_list attributes;

    /* Hide screen gridlines if required */
    if (!screen_gridlines)
        attributes.push_back({"showGridLines", "0"});

    /* Hide zeroes in cells. */
    if (!show_zeros) {
        attributes.push_back({"showZeros", "0"});
    }

    /* Display worksheet right to left for Hebrew, Arabic and others. */
    if (right_to_left) {
        attributes.push_back({"rightToLeft", "1"});
    }

    /* Show that the sheet tab is selected. */
    if (selected)
        attributes.push_back({"tabSelected", "1"});

    /* Turn outlines off. Also required in the outlinePr element. */
    if (!outline_on) {
        attributes.push_back({"showOutlineSymbols", "0"});
    }

    /* Set the page view/layout mode if required. */
    if (page_view)
        attributes.push_back({"view", "pageLayout"});

    /* Set the zoom level. */
    if (zoom != 100) {
        if (!page_view) {
            attributes.push_back({"zoomScale", std::to_string(zoom)});

            if (zoom_scale_normal)
                attributes.push_back({"zoomScaleNormal", std::to_string(zoom)});
        }
    }

    attributes.push_back({"workbookViewId", "0"});

    if (panes.type != NO_PANES || !selections.empty()) {
        lxw_xml_start_tag("sheetView", attributes);
       _write_panes();
       _write_selections();
        lxw_xml_end_tag("sheetView");
    }
    else {
        lxw_xml_empty_tag("sheetView", attributes);
    }
}

/*
 * Write the <sheetViews> element.
 */
void
worksheet::_write_sheet_views()
{
    lxw_xml_start_tag("sheetViews");

    /* Write the sheetView element. */
   _write_sheet_view();

    lxw_xml_end_tag("sheetViews");
}

/*
 * Write the <sheetFormatPr> element.
 */
void worksheet::_write_sheet_format_pr()
{
    xml_attribute_list attributes = {
        {"defaultRowHeight", std::to_string(default_row_height)}
    };

    if (default_row_height != LXW_DEF_ROW_HEIGHT)
        attributes.push_back({"customHeight", "1"});

    if (default_row_zeroed)
        attributes.push_back({"zeroHeight", "1"});

    lxw_xml_empty_tag("sheetFormatPr", attributes);
}

/*
 * Write the <sheetData> element.
 */
void
worksheet::_write_sheet_data()
{
    if (RB_EMPTY(table)) {
        lxw_xml_empty_tag("sheetData");
    }
    else {
        lxw_xml_start_tag("sheetData");
       _write_rows();
        lxw_xml_end_tag("sheetData");
    }
}

/*
 * Write the <sheetData> element when the memory optimization is on. In which
 * case we read the data stored in the temp file and rewrite it to the XML
 * sheet file.
 */
void worksheet::_write_optimized_sheet_data()
{
    size_t read_size = 1;
    char buffer[LXW_BUFFER_SIZE];

    if (dim_rowmin == LXW_ROW_MAX) {
        /* If the dimensions aren't defined then there is no data to write. */
        lxw_xml_empty_tag("sheetData");
    }
    else {

        lxw_xml_start_tag("sheetData");

        /* Flush and rewind the temp file. */
        fflush(optimize_tmpfile);
        rewind(optimize_tmpfile);

        while (read_size) {
            read_size =
                fread(buffer, 1, LXW_BUFFER_SIZE, optimize_tmpfile);
            fwrite(buffer, 1, read_size, file);
        }

        fclose(optimize_tmpfile);

        lxw_xml_end_tag("sheetData");
    }
}

/*
 * Write the <pageMargins> element.
 */
void worksheet::_write_page_margins()
{
    xml_attribute_list attributes = {
        {"left", std::to_string(margin_left)},
        {"right", std::to_string(margin_right)},
        {"top", std::to_string(margin_top)},
        {"bottom", std::to_string(margin_bottom)},
        {"header", std::to_string(margin_header)},
        {"footer", footer}
    };
    lxw_xml_empty_tag("pageMargins", attributes);
}

/*
 * Write the <pageSetup> element.
 * The following is an example taken from Excel.
 * <pageSetup
 *     paperSize="9"
 *     scale="110"
 *     fitToWidth="2"
 *     fitToHeight="2"
 *     pageOrder="overThenDown"
 *     orientation="portrait"
 *     blackAndWhite="1"
 *     draft="1"
 *     horizontalDpi="200"
 *     verticalDpi="200"
 *     r:id="rId1"
 * />
 */
void
worksheet::_write_page_setup()
{
    xml_attribute_list attributes;

    if (!page_setup_changed)
        return;

    /* Set paper size. */
    if (paper_size)
        attributes.push_back({"paperSize", std::to_string(paper_size)});

    /* Set the print_scale. */
    if (print_scale != 100)
        attributes.push_back({"scale", std::to_string(print_scale)});

    /* Set the "Fit to page" properties. */
    if (fit_page && fit_width != 1)
        attributes.push_back({"fitToWidth", std::to_string(fit_width)});

    if (fit_page && fit_height != 1)
        attributes.push_back({"fitToHeight", std::to_string(fit_height)});

    /* Set the page print direction. */
    if (page_order)
        attributes.push_back({"pageOrder", "overThenDown"});

    /* Set start page. */
    if (page_start > 1)
        attributes.push_back({"firstPageNumber", std::to_string(page_start)});

    /* Set page orientation. */
    if (orientation)
        attributes.push_back({"orientation", "portrait"});
    else
        attributes.push_back({"orientation", "landscape"});

    /* Set start page active flag. */
    if (page_start)
        attributes.push_back({"useFirstPageNumber", "1"});

    /* Set the DPI. Mainly only for testing. */
    if (horizontal_dpi)
        attributes.push_back({"horizontalDpi", std::to_string(horizontal_dpi)});

    if (vertical_dpi)
        attributes.push_back({"verticalDpi", std::to_string(vertical_dpi)});

    lxw_xml_empty_tag("pageSetup", attributes);
}

/*
 * Write the <printOptions> element.
 */
void
worksheet::_write_print_options()
{
    xml_attribute_list attributes;

    if (!print_options_changed)
        return;

    /* Set horizontal centering. */
    if (hcenter) {
        attributes.push_back({"horizontalCentered", "1"});
    }

    /* Set vertical centering. */
    if (vcenter) {
        attributes.push_back({"verticalCentered", "1"});
    }

    /* Enable row and column headers. */
    if (print_headers) {
        attributes.push_back({"headings", "1"});
    }

    /* Set printed gridlines. */
    if (print_gridlines) {
        attributes.push_back({"gridLines", "1"});
    }

    lxw_xml_empty_tag("printOptions", attributes);


}

/*
 * Write the <row> element.
 */
void worksheet::_write_row(lxw_row *row, const std::string& spans)
{
    xml_attribute_list attributes;

    int32_t xf_index = 0;
    double height;

    if (row->format) {
        xf_index = row->format->get_xf_index();
    }

    if (row->height_changed)
        height = row->height;
    else
        height = default_row_height;

    attributes.push_back({"r", std::to_string(row->row_num + 1)});

    if (!spans.empty())
        attributes.push_back({"spans", spans});

    if (xf_index)
        attributes.push_back({"s", std::to_string(xf_index)});

    if (row->format)
        attributes.push_back({"customFormat", "1"});

    if (height != LXW_DEF_ROW_HEIGHT)
        attributes.push_back({"ht", std::to_string(height)});

    if (row->hidden)
        attributes.push_back({"hidden", "1"});

    if (height != LXW_DEF_ROW_HEIGHT)
        attributes.push_back({"customHeight", "1"});

    if (row->collapsed)
        attributes.push_back({"collapsed", "1"});

    if (!row->data_changed)
        lxw_xml_empty_tag("row", attributes);
    else
        lxw_xml_start_tag("row", attributes);


}

/*
 * Convert the width of a cell from user's units to pixels. Excel rounds the
 * column width to the nearest pixel. If the width hasn't been set by the user
 * we use the default value. If the column is hidden it has a value of zero.
 */
int32_t worksheet::_size_col(lxw_col_t col_num)
{
    lxw_col_options *col_opt = NULL;
    uint32_t pixels;
    double width;
    double max_digit_width = 7.0;       /* For Calabri 11. */
    double padding = 5.0;
    lxw_col_t col_index;

    /* Search for the col number in the array of col_options. Each col_option
     * entry contains the start and end column for a range.
     */
    for (col_index = 0; col_index < col_options_max; col_index++) {
        col_opt = col_options[col_index];

        if (col_opt) {
            if (col_num >= col_opt->firstcol && col_num <= col_opt->lastcol)
                break;
            else
                col_opt = NULL;
        }
    }

    if (col_opt) {
        width = col_opt->width;

        /* Convert to pixels. */
        if (width == 0) {
            pixels = 0;
        }
        else if (width < 1.0) {
            pixels = (uint32_t) (width * (max_digit_width + padding) + 0.5);
        }
        else {
            pixels = (uint32_t) (width * max_digit_width + 0.5) + 5;
        }
    }
    else {
        pixels = default_col_pixels;
    }

    return pixels;
}

/*
 * Convert the height of a cell from user's units to pixels. If the height
 * hasn't been set by the user we use the default value. If the row is hidden
 * it has a value of zero.
 */
int32_t worksheet::_size_row(lxw_row_t row_num)
{
    lxw_row *row;
    uint32_t pixels;
    double height;

    row = find_row(row_num);

    if (row) {
        height = row->height;

        if (height == 0)
            pixels = 0;
        else
            pixels = (uint32_t) (4.0 / 3.0 * height);
    }
    else {
        pixels = (uint32_t) (4.0 / 3.0 * default_row_height);
    }

    return pixels;
}

/*
 * Calculate the vertices that define the position of a graphical object
 * within the worksheet in pixels.
 *         +------------+------------+
 *         |     A      |      B     |
 *   +-----+------------+------------+
 *   |     |(x1,y1)     |            |
 *   |  1  |(A1)._______|______      |
 *   |     |    |              |     |
 *   |     |    |              |     |
 *   +-----+----|    BITMAP    |-----+
 *   |     |    |              |     |
 *   |  2  |    |______________.     |
 *   |     |            |        (B2)|
 *   |     |            |     (x2,y2)|
 *   +---- +------------+------------+
 *
 * Example of an object that covers some of the area from cell A1 to cell B2.
 * Based on the width and height of the object we need to calculate 8 vars:
 *
 *     col_start, row_start, col_end, row_end, x1, y1, x2, y2.
 *
 * We also calculate the absolute x and y position of the top left vertex of
 * the object. This is required for images:
 *
 *    x_abs, y_abs
 *
 * The width and height of the cells that the object occupies can be variable
 * and have to be taken into account.
 *
 * The values of col_start and row_start are passed in from the calling
 * function. The values of col_end and row_end are calculated by subtracting
 * the width and height of the object from the width and height of the
 * underlying cells.
 */
void worksheet::_position_object_pixels(const image_options_ptr& image,
                                  const drawing_object_ptr& drawing_object)
{
    lxw_col_t col_start;        /* Column containing upper left corner.  */
    int32_t x1;                 /* Distance to left side of object.      */

    lxw_row_t row_start;        /* Row containing top left corner.       */
    int32_t y1;                 /* Distance to top of object.            */

    lxw_col_t col_end;          /* Column containing lower right corner. */
    double x2;                  /* Distance to right side of object.     */

    lxw_row_t row_end;          /* Row containing bottom right corner.   */
    double y2;                  /* Distance to bottom of object.         */

    double width;               /* Width of object frame.                */
    double height;              /* Height of object frame.               */

    uint32_t x_abs = 0;         /* Abs. distance to left side of object. */
    uint32_t y_abs = 0;         /* Abs. distance to top  side of object. */

    uint32_t i;

    col_start = image->col;
    row_start = image->row;
    x1 = image->x_offset;
    y1 = image->y_offset;
    width = image->width;
    height = image->height;

    /* Adjust start column for negative offsets. */
    while (x1 < 0 && col_start > 0) {
        x1 += _size_col(col_start - 1);
        col_start--;
    }

    /* Adjust start row for negative offsets. */
    while (y1 < 0 && row_start > 0) {
        y1 += _size_row(row_start - 1);
        row_start--;
    }

    /* Ensure that the image isn't shifted off the page at top left. */
    if (x1 < 0)
        x1 = 0;

    if (y1 < 0)
        y1 = 0;

    /* Calculate the absolute x offset of the top-left vertex. */
    if (col_size_changed) {
        for (i = 0; i < col_start; i++)
            x_abs += _size_col(i);
    }
    else {
        /* Optimization for when the column widths haven't changed. */
        x_abs += default_col_pixels * col_start;
    }

    x_abs += x1;

    /* Calculate the absolute y offset of the top-left vertex. */
    /* Store the column change to allow optimizations. */
    if (row_size_changed) {
        for (i = 0; i < row_start; i++)
            y_abs += _size_row(i);
    }
    else {
        /* Optimization for when the row heights haven"t changed. */
        y_abs += default_row_pixels * row_start;
    }

    y_abs += y1;

    /* Adjust start col for offsets that are greater than the col width. */
    while (x1 >= _size_col(col_start)) {
        x1 -= _size_col(col_start);
        col_start++;
    }

    /* Adjust start row for offsets that are greater than the row height. */
    while (y1 >= _size_row(row_start)) {
        y1 -= _size_row(row_start);
        row_start++;
    }

    /* Initialize end cell to the same as the start cell. */
    col_end = col_start;
    row_end = row_start;

    width = width + x1;
    height = height + y1;

    /* Subtract the underlying cell widths to find the end cell. */
    while (width >= _size_col(col_end)) {
        width -= _size_col(col_end);
        col_end++;
    }

    /* Subtract the underlying cell heights to find the end cell. */
    while (height >= _size_row(row_end)) {
        height -= _size_row(row_end);
        row_end++;
    }

    /* The end vertices are whatever is left from the width and height. */
    x2 = width;
    y2 = height;

    /* Add the dimensions to the drawing object. */
    drawing_object->from.col = col_start;
    drawing_object->from.row = row_start;
    drawing_object->from.col_offset = x1;
    drawing_object->from.row_offset = y1;
    drawing_object->to.col = col_end;
    drawing_object->to.row = row_end;
    drawing_object->to.col_offset = x2;
    drawing_object->to.row_offset = y2;
    drawing_object->col_absolute = x_abs;
    drawing_object->row_absolute = y_abs;

}

/*
 * Calculate the vertices that define the position of a graphical object
 * within the worksheet in EMUs. The vertices are expressed as English
 * Metric Units (EMUs). There are 12,700 EMUs per point.
 * Therefore, 12,700 * 3 /4 = 9,525 EMUs per pixel.
 */
void worksheet::_position_object_emus( const image_options_ptr& image,
                                 const drawing_object_ptr& drawing_object)
{

    _position_object_pixels(image, drawing_object);

    /* Convert the pixel values to EMUs. See above. */
    drawing_object->from.col_offset *= 9525;
    drawing_object->from.row_offset *= 9525;
    drawing_object->to.col_offset *= 9525;
    drawing_object->to.row_offset *= 9525;
    drawing_object->to.col_offset += 0.5;
    drawing_object->to.row_offset += 0.5;
    drawing_object->col_absolute *= 9525;
    drawing_object->row_absolute *= 9525;
}

/*
 * Set up image/drawings.
 */
void worksheet::prepare_image(uint16_t image_ref_id, uint16_t drawing_id,
                            const image_options_ptr& image_data)
{
    drawing_object_ptr drawing_object;
    rel_tuple_ptr relationship;
    double width;
    double height;
    char filename[LXW_FILENAME_LENGTH];

    if (!drawing) {
        drawing = std::make_shared<xlsxwriter::drawing>();
        drawing->embedded = true;

        relationship = std::make_shared<xlsxwriter::rel_tuple>();

        relationship->type = "/drawing";

        lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                     "../drawings/drawing%d.xml", drawing_id);

        relationship->target = filename;

        external_drawing_links.push_back(relationship);
    }

    drawing_object = std::make_shared<xlsxwriter::drawing_object>();

    drawing_object->anchor_type = LXW_ANCHOR_TYPE_IMAGE;
    drawing_object->edit_as = LXW_ANCHOR_EDIT_AS_ONE_CELL;
    drawing_object->description = image_data->short_name;

    /* Scale to user scale. */
    width = image_data->width * image_data->x_scale;
    height = image_data->height * image_data->y_scale;

    /* Scale by non 96dpi resolutions. */
    width *= 96.0 / image_data->x_dpi;
    height *= 96.0 / image_data->y_dpi;

    /* Convert to the nearest pixel. */
    image_data->width = width;
    image_data->height = height;

    _position_object_emus(image_data, drawing_object);

    /* Convert from pixels to emus. */
    drawing_object->width = (uint32_t) (0.5 + width * 9525);
    drawing_object->height = (uint32_t) (0.5 + height * 9525);

    drawing->add_drawing_object(drawing_object);

    relationship = std::make_shared<xlsxwriter::rel_tuple>();

    relationship->type = "/image";

    lxw_snprintf(filename, 32, "../media/image%d.%s", image_ref_id,
                 image_data->extension);

    relationship->target = filename;

    drawing_links.push_front(relationship);
}

/*
 * Set up chart/drawings.
 */
void worksheet::prepare_chart(
                            uint16_t chart_ref_id, uint16_t drawing_id,
                            const image_options_ptr& image_data)
{
    drawing_object_ptr drawing_object;
    rel_tuple_ptr relationship;
    double width;
    double height;
    char filename[LXW_FILENAME_LENGTH];

    if (!drawing) {
        drawing = std::make_shared<xlsxwriter::drawing>();
        drawing->embedded = true;

        relationship = std::make_shared<rel_tuple>();

        relationship->type = "/drawing";

        lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                     "../drawings/drawing%d.xml", drawing_id);

        relationship->target = filename;
        external_drawing_links.push_back(relationship);
    }

    drawing_object = std::make_shared<xlsxwriter::drawing_object>();

    drawing_object->anchor_type = LXW_ANCHOR_TYPE_CHART;
	switch (image_data->anchor)
	{
	case 1:
		drawing_object->edit_as = LXW_ANCHOR_EDIT_AS_RELATIVE;
		break;
	case 2:
		drawing_object->edit_as = LXW_ANCHOR_EDIT_AS_ONE_CELL;
		break;
	case 3:
		drawing_object->edit_as = LXW_ANCHOR_EDIT_AS_ABSOLUTE;
		break;
	default:
		drawing_object->edit_as = LXW_ANCHOR_EDIT_AS_ONE_CELL;
		break;
	}
    drawing_object->description = "TODO_DESC";

    /* Scale to user scale. */
    width = image_data->width * image_data->x_scale;
    height = image_data->height * image_data->y_scale;

    /* Convert to the nearest pixel. */
    image_data->width = width;
    image_data->height = height;

    _position_object_emus(image_data, drawing_object);

    /* Convert from pixels to emus. */
    drawing_object->width = (uint32_t) (0.5 + width * 9525);
    drawing_object->height = (uint32_t) (0.5 + height * 9525);

    drawing->add_drawing_object(drawing_object);

    relationship = std::make_shared<rel_tuple>();

    relationship->type = "/chart";

    lxw_snprintf(filename, 32, "../charts/chart%d.xml", chart_ref_id);

    relationship->target = filename;

    drawing_links.push_back(relationship);
}

/*
 * Extract width and height information from a PNG file.
 */
lxw_error
_process_png(const image_options_ptr& image_options)
{
    uint32_t length;
    uint32_t offset;
    char type[4];
    uint32_t width = 0;
    uint32_t height = 0;
    double x_dpi = 96;
    double y_dpi = 96;

    FILE *stream = image_options->stream;

    /* Skip another 4 bytes to the end of the PNG header. */
    fseek(stream, 4, SEEK_CUR);

    while (!feof(stream)) {

        /* Read the PNG length and type fields for the sub-section. */
        if (fread(&length, sizeof(length), 1, stream) < 1)
            break;

        if (fread(&type, 1, 4, stream) < 4)
            break;

        /* Convert the length to network order. */
        length = LXW_UINT32_NETWORK(length);

        /* The offset for next fseek() is the field length + type length. */
        offset = length + 4;

        if (memcmp(type, "IHDR", 4) == 0) {
            if (fread(&width, sizeof(width), 1, stream) < 1)
                break;

            if (fread(&height, sizeof(height), 1, stream) < 1)
                break;

            width = LXW_UINT32_NETWORK(width);
            height = LXW_UINT32_NETWORK(height);

            /* Reduce the offset by the length of previous freads(). */
            offset -= 8;
        }

        if (memcmp(type, "pHYs", 4) == 0) {
            uint32_t x_ppu = 0;
            uint32_t y_ppu = 0;
            uint8_t units = 1;

            if (fread(&x_ppu, sizeof(x_ppu), 1, stream) < 1)
                break;

            if (fread(&y_ppu, sizeof(y_ppu), 1, stream) < 1)
                break;

            if (fread(&units, sizeof(units), 1, stream) < 1)
                break;

            if (units == 1) {
                x_ppu = LXW_UINT32_NETWORK(x_ppu);
                y_ppu = LXW_UINT32_NETWORK(y_ppu);

                x_dpi = (double) x_ppu *0.0254;
                y_dpi = (double) y_ppu *0.0254;
            }

            /* Reduce the offset by the length of previous freads(). */
            offset -= 9;
        }

        if (memcmp(type, "IEND", 4) == 0)
            break;

        if (!feof(stream))
            fseek(stream, offset, SEEK_CUR);
    }

    /* Ensure that we read some valid data from the file. */
    if (width == 0) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "no size data found in file: %s.",
                         image_options->filename);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }

    /* Set the image metadata. */
    image_options->image_type = LXW_IMAGE_PNG;
    image_options->width = width;
    image_options->height = height;
    image_options->x_dpi = x_dpi ? x_dpi : 96;
    image_options->y_dpi = y_dpi ? x_dpi : 96;
    image_options->extension = "png";

    return LXW_NO_ERROR;
}

/*
 * Extract width and height information from a JPEG file.
 */
lxw_error
_process_jpeg(const image_options_ptr& image_options)
{
    uint16_t length;
    uint16_t marker;
    uint32_t offset;
    uint16_t width = 0;
    uint16_t height = 0;
    double x_dpi = 96;
    double y_dpi = 96;

    FILE *stream = image_options->stream;

    /* Read back 2 bytes to the end of the initial 0xFFD8 marker. */
    fseek(stream, -2, SEEK_CUR);

    /* Search through the image data to read the height and width in the */
    /* 0xFFC0/C2 element. Also read the DPI in the 0xFFE0 element. */
    while (!feof(stream)) {

        /* Read the JPEG marker and length fields for the sub-section. */
        if (fread(&marker, sizeof(marker), 1, stream) < 1)
            break;

        if (fread(&length, sizeof(length), 1, stream) < 1)
            break;

        /* Convert the marker and length to network order. */
        marker = LXW_UINT16_NETWORK(marker);
        length = LXW_UINT16_NETWORK(length);

        /* The offset for next fseek() is the field length + type length. */
        offset = length - 2;

        if (marker == 0xFFC0 || marker == 0xFFC2) {
            /* Skip 1 byte to height and width. */
            fseek(stream, 1, SEEK_CUR);

            if (fread(&height, sizeof(height), 1, stream) < 1)
                break;

            if (fread(&width, sizeof(width), 1, stream) < 1)
                break;

            height = LXW_UINT16_NETWORK(height);
            width = LXW_UINT16_NETWORK(width);

            offset -= 9;
        }

        if (marker == 0xFFE0) {
            uint16_t x_density = 0;
            uint16_t y_density = 0;
            uint8_t units = 1;

            fseek(stream, 7, SEEK_CUR);

            if (fread(&units, sizeof(units), 1, stream) < 1)
                break;

            if (fread(&x_density, sizeof(x_density), 1, stream) < 1)
                break;

            if (fread(&y_density, sizeof(y_density), 1, stream) < 1)
                break;

            x_density = LXW_UINT16_NETWORK(x_density);
            y_density = LXW_UINT16_NETWORK(y_density);

            if (units == 1) {
                x_dpi = x_density;
                y_dpi = y_density;
            }

            if (units == 2) {
                x_dpi = x_density * 2.54;
                y_dpi = y_density * 2.54;
            }

            offset -= 12;
        }

        if (marker == 0xFFDA)
            break;

        if (!feof(stream))
            fseek(stream, offset, SEEK_CUR);
    }

    /* Ensure that we read some valid data from the file. */
    if (width == 0) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "no size data found in file: %s.",
                         image_options->filename);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }

    /* Set the image metadata. */
    image_options->image_type = LXW_IMAGE_JPEG;
    image_options->width = width;
    image_options->height = height;
    image_options->x_dpi = x_dpi ? x_dpi : 96;
    image_options->y_dpi = y_dpi ? x_dpi : 96;
    image_options->extension = "jpeg";

    return LXW_NO_ERROR;
}

/*
 * Extract width and height information from a BMP file.
 */
lxw_error
_process_bmp(const image_options_ptr& image_options)
{
    uint32_t width = 0;
    uint32_t height = 0;
    double x_dpi = 96;
    double y_dpi = 96;

    FILE *stream = image_options->stream;

    /* Skip another 14 bytes to the start of the BMP height/width. */
    fseek(stream, 14, SEEK_CUR);

    if (fread(&width, sizeof(width), 1, stream) < 1)
        width = 0;

    if (fread(&height, sizeof(height), 1, stream) < 1)
        height = 0;

    /* Ensure that we read some valid data from the file. */
    if (width == 0) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "no size data found in file: %s.",
                         image_options->filename);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }

    /* Set the image metadata. */
    image_options->image_type = LXW_IMAGE_BMP;
    image_options->width = width;
    image_options->height = height;
    image_options->x_dpi = x_dpi;
    image_options->y_dpi = y_dpi;
    image_options->extension = "bmp";

    return LXW_NO_ERROR;
}

/*
 * Extract information from the image file such as dimension, type, filename,
 * and extension.
 */
lxw_error _get_image_properties(const image_options_ptr& image_options)
{
    unsigned char signature[4];

    /* Read 4 bytes to look for the file header/signature. */
    if (fread(signature, 1, 4, image_options->stream) < 4) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "couldn't read file type for file: %s.",
                         image_options->filename);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }

    if (memcmp(&signature[1], "PNG", 3) == 0) {
        if (_process_png(image_options) != LXW_NO_ERROR)
            return LXW_ERROR_IMAGE_DIMENSIONS;
    }
    else if (signature[0] == 0xFF && signature[1] == 0xD8) {
        if (_process_jpeg(image_options) != LXW_NO_ERROR)
            return LXW_ERROR_IMAGE_DIMENSIONS;
    }
    else if (memcmp(signature, "BM", 2) == 0) {
        if (_process_bmp(image_options) != LXW_NO_ERROR)
            return LXW_ERROR_IMAGE_DIMENSIONS;
    }
    else {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "unsupported image format for file: %s.",
                         image_options->filename);
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }

    return LXW_NO_ERROR;
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Write out a number worksheet cell. Doesn't use the xml functions as an
 * optimization in the inner cell writing loop.
 */
void worksheet::_write_number_cell(const std::string& range,
                   int32_t style_index, lxw_cell *cell)
{
    if (style_index)
        fprintf(file,
                "<c r=\"%s\" s=\"%d\"><v>%.16g</v></c>",
                range.c_str(), style_index, cell->u.number);
    else
        fprintf(file,
                "<c r=\"%s\"><v>%.16g</v></c>", range.c_str(), cell->u.number);
}

/*
 * Write out a string worksheet cell. Doesn't use the xml functions as an
 * optimization in the inner cell writing loop.
 */
void worksheet::_write_string_cell(const std::string& range,
                   int32_t style_index, lxw_cell *cell)
{
    if (style_index)
        fprintf(file,
                "<c r=\"%s\" s=\"%d\" t=\"s\"><v>%d</v></c>",
                range.c_str(), style_index, cell->u.string_id);
    else
        fprintf(file,
                "<c r=\"%s\" t=\"s\"><v>%d</v></c>",
                range.c_str(), cell->u.string_id);
}

/*
 * Write out an inline string. Doesn't use the xml functions as an
 * optimization in the inner cell writing loop.
 */
void worksheet::_write_inline_string_cell(const std::string& range,
                          int32_t style_index, lxw_cell *cell)
{
    std::string string = lxw_escape_data(*cell->u.string);

    /* Add attribute to preserve leading or trailing whitespace. */
    if (isspace(string[0])
        || isspace(string[string.size() - 1])) {

        if (style_index)
            fprintf(file,
                    "<c r=\"%s\" s=\"%d\" t=\"inlineStr\"><is>"
                    "<t xml:space=\"preserve\">%s</t></is></c>",
                    range.c_str(), style_index, string.c_str());
        else
            fprintf(file,
                    "<c r=\"%s\" t=\"inlineStr\"><is>"
                    "<t xml:space=\"preserve\">%s</t></is></c>",
                    range.c_str(), string.c_str());
    }
    else {
        if (style_index)
            fprintf(file,
                    "<c r=\"%s\" s=\"%d\" t=\"inlineStr\">"
                    "<is><t>%s</t></is></c>", range.c_str(), style_index, string.c_str());
        else
            fprintf(file,
                    "<c r=\"%s\" t=\"inlineStr\">"
                    "<is><t>%s</t></is></c>", range.c_str(), string.c_str());
    }
}

/*
 * Write out a formula worksheet cell with a numeric result.
 */
void worksheet::_write_formula_num_cell(lxw_cell *cell)
{
    char data[LXW_ATTR_32];

    lxw_snprintf(data, LXW_ATTR_32, "%.16g", cell->formula_result);

    lxw_xml_data_element("f", *cell->u.string);
    lxw_xml_data_element("v", data);
}

/*
 * Write out an array formula worksheet cell with a numeric result.
 */
void worksheet::_write_array_formula_num_cell(lxw_cell *cell)
{
    char data[LXW_ATTR_32];

    xml_attribute_list attributes = {
        {"t", "array"},
        {"ref", *cell->user_data1}
    };

    lxw_snprintf(data, LXW_ATTR_32, "%.16g", cell->formula_result);

    lxw_xml_data_element("f", *cell->u.string, attributes);
    lxw_xml_data_element("v", data);


}

/*
 * Write out a boolean worksheet cell.
 */
void worksheet::_write_boolean_cell(lxw_cell *cell)
{
    std::string data("0");

    if (cell->u.number)
        data[0] = '1';
    else
        data[0] = '0';

    lxw_xml_data_element("v", data);
}

/*
 * Calculate the "spans" attribute of the <row> tag. This is an XLSX
 * optimization and isn't strictly required. However, it makes comparing
 * files easier.
 *
 * The span is the same for each block of 16 rows.
 */
void
_calculate_spans(struct lxw_row *row, char *span, int32_t *block_num)
{
    lxw_col_t span_col_min = RB_MIN(lxw_table_cells, row->cells)->col_num;
    lxw_col_t span_col_max = RB_MAX(lxw_table_cells, row->cells)->col_num;
    lxw_col_t col_min;
    lxw_col_t col_max;
    *block_num = row->row_num / 16;

    row = RB_NEXT(lxw_table_rows, root, row);

    while (row && (int32_t) (row->row_num / 16) == *block_num) {

        if (!RB_EMPTY(row->cells)) {
            col_min = RB_MIN(lxw_table_cells, row->cells)->col_num;
            col_max = RB_MAX(lxw_table_cells, row->cells)->col_num;

            if (col_min < span_col_min)
                span_col_min = col_min;

            if (col_max > span_col_max)
                span_col_max = col_max;
        }

        row = RB_NEXT(lxw_table_rows, root, row);
    }

    lxw_snprintf(span, LXW_MAX_CELL_RANGE_LENGTH,
                 "%d:%d", span_col_min + 1, span_col_max + 1);
}

/*
 * Write out a generic worksheet cell.
 */
void worksheet::_write_cell(lxw_cell *cell, xlsxwriter::format* row_format)
{
    std::string range;
    lxw_row_t row_num = cell->row_num;
    lxw_col_t col_num = cell->col_num;
    int32_t style_index = 0;

    lxw_rowcol_to_cell(range, row_num, col_num);

    if (cell->format) {
        style_index = cell->format->get_xf_index();
    }
    else if (row_format) {
        style_index = row_format->get_xf_index();
    }
    else if (col_num < col_formats_max && col_formats[col_num]) {
        style_index = col_formats[col_num]->get_xf_index();
    }

    /* Unrolled optimization for most commonly written cell types. */
    if (cell->type == NUMBER_CELL) {
        _write_number_cell(range, style_index, cell);
        return;
    }

    if (cell->type == STRING_CELL) {
        _write_string_cell(range, style_index, cell);
        return;
    }

    if (cell->type == INLINE_STRING_CELL) {
        _write_inline_string_cell(range, style_index, cell);
        return;
    }

    /* For other cell types use the general functions. */
    xml_attribute_list attributes = {
        {"r", range}
    };

    if (style_index)
        attributes.push_back({"s", std::to_string(style_index)});

    if (cell->type == FORMULA_CELL) {
        lxw_xml_start_tag("c", attributes);
        _write_formula_num_cell(cell);
        lxw_xml_end_tag("c");
    }
    else if (cell->type == BLANK_CELL) {
        lxw_xml_empty_tag("c", attributes);
    }
    else if (cell->type == BOOLEAN_CELL) {
        attributes.push_back({"t", "b"});
        lxw_xml_start_tag("c", attributes);
        _write_boolean_cell(cell);
        lxw_xml_end_tag("c");
    }
    else if (cell->type == ARRAY_FORMULA_CELL) {
        lxw_xml_start_tag("c", attributes);
        _write_array_formula_num_cell(cell);
        lxw_xml_end_tag("c");
    }
}

/*
 * Write out the worksheet data as a series of rows and cells.
 */
void worksheet::_write_rows()
{
    lxw_row *row;
    lxw_cell *cell;
    int32_t block_num = -1;
    char spans[LXW_MAX_CELL_RANGE_LENGTH] = { 0 };

    RB_FOREACH(row, lxw_table_rows, table) {

        if (RB_EMPTY(row->cells)) {
            /* Row contains no cells but has height, format or other data. */

            /* Write a default span for default rows. */
            if (default_row_set)
                _write_row(row, "1:1");
            else
                _write_row(row, NULL);
        }
        else {
            /* Row and cell data. */
            if ((int32_t) row->row_num / 16 > block_num)
                _calculate_spans(row, spans, &block_num);

            _write_row(row, spans);

            RB_FOREACH(cell, lxw_table_cells, row->cells) {
                _write_cell(cell, row->format);
            }
            lxw_xml_end_tag("row");
        }
    }
}

/*
 * Write out the worksheet data as a single row with cells. This method is
 * used when memory optimization is on. A single row is written and the data
 * array is reset. That way only one row of data is kept in memory at any one
 * time. We don't write span data in the optimized case since it is optional.
 */
void worksheet::write_single_row()
{
    lxw_row *row = optimize_row;
    lxw_col_t col;

    /* skip row if it doesn't contain row formatting, cell data or a comment. */
    if (!(row->row_changed || row->data_changed))
        return;

    /* Write the cells if the row contains data. */
    if (!row->data_changed) {
        /* Row data only. No cells. */
        _write_row(row, NULL);
    }
    else {
        /* Row and cell data. */
        _write_row(row, NULL);

        for (col = dim_colmin; col <= dim_colmax; col++) {
            if (array[col]) {
                _write_cell(array[col], row->format);
                _free_cell(array[col]);
                array[col] = NULL;
            }
        }

        lxw_xml_end_tag("row");
    }

    /* Reset the row. */
    row->height = LXW_DEF_ROW_HEIGHT;
    row->format = NULL;
    row->hidden = false;
    row->level = 0;
    row->collapsed = false;
    row->data_changed = false;
    row->row_changed = false;
}

/*
 * Write the <col> element.
 */
void worksheet::_write_col_info(lxw_col_options *options)
{
    double width = options->width;
    uint8_t has_custom_width = true;
    int32_t xf_index = 0;
    double max_digit_width = 7.0;       /* For Calabri 11. */
    double padding = 5.0;

    /* Get the format index. */
    if (options->format) {
        xf_index = options->format->get_xf_index();
    }

    /* Check if width is the Excel default. */
    if (width == LXW_DEF_COL_WIDTH) {

        /* The default col width changes to 0 for hidden columns. */
        if (options->hidden)
            width = 0;
        else
            has_custom_width = false;

    }

    /* Convert column width from user units to character width. */
    if (width > 0) {
        if (width < 1) {
            width = (uint16_t) (((uint16_t)
                                 (width * (max_digit_width + padding) + 0.5))
                                / max_digit_width * 256.0) / 256.0;
        }
        else {
            width = (uint16_t) (((uint16_t)
                                 (width * max_digit_width + 0.5) + padding)
                                / max_digit_width * 256.0) / 256.0;
        }
    }

    xml_attribute_list attributes = {
        {"min", std::to_string(1 + options->firstcol)},
        {"max", std::to_string(1 + options->lastcol)},
        {"width", std::to_string(width)}
    };

    if (xf_index)
        attributes.push_back({"style", std::to_string(xf_index)});

    if (options->hidden)
        attributes.push_back({"hidden", "1"});

    if (has_custom_width)
        attributes.push_back({"customWidth", "1"});

    if (options->level)
        attributes.push_back({"outlineLevel", std::to_string(options->level)});

    if (options->collapsed)
        attributes.push_back({"collapsed", "1"});

    lxw_xml_empty_tag("col", attributes);


}

/*
 * Write the <cols> element and <col> sub elements.
 */
void worksheet::_write_cols()
{
    lxw_col_t col;

    if (!col_size_changed)
        return;

    lxw_xml_start_tag("cols");

    for (col = 0; col < col_options_max; col++) {
        if (col_options[col])
           _write_col_info(col_options[col]);
    }

    lxw_xml_end_tag("cols");
}

/*
 * Write the <mergeCell> element.
 */
void worksheet::_write_merge_cell(const std::shared_ptr<lxw_merged_range>& merged_range)
{
    std::string ref;

    /* Convert the merge dimensions to a cell range. */
    lxw_rowcol_to_range(ref, merged_range->first_row, merged_range->first_col,
                        merged_range->last_row, merged_range->last_col);

    xml_attribute_list attributes = {
        {"ref", ref}
    };

    lxw_xml_empty_tag("mergeCell", attributes);
}

/*
 * Write the <mergeCells> element.
 */
void worksheet::_write_merge_cells()
{
    if (merged_range_count) {
        xml_attribute_list attributes = {
            {"count", std::to_string(merged_range_count)}
        };

        lxw_xml_start_tag("mergeCells", attributes);

        for (const auto& merged_range: merged_ranges) {
           _write_merge_cell(merged_range);
        }
        lxw_xml_end_tag("mergeCells");
    }
}

/*
 * Write the <oddHeader> element.
 */
void worksheet::_write_odd_header()
{
    lxw_xml_data_element("oddHeader", header);
}

/*
 * Write the <oddFooter> element.
 */
void
worksheet::_write_odd_footer()
{
    lxw_xml_data_element("oddFooter", footer);
}

/*
 * Write the <headerFooter> element.
 */
void
worksheet::_write_header_footer()
{
    if (!header_footer_changed)
        return;

    lxw_xml_start_tag("headerFooter");

    if (header[0] != '\0')
       _write_odd_header();

    if (footer[0] != '\0')
       _write_odd_footer();

    lxw_xml_end_tag("headerFooter");
}

/*
 * Write the <pageSetUpPr> element.
 */
void worksheet::_write_page_set_up_pr()
{
    if (!fit_page)
        return;

    xml_attribute_list attributes = {
        {"fitToPage", "1"}
    };

    lxw_xml_empty_tag("pageSetUpPr", attributes);
}

/*
 * Write the <tabColor> element.
 */
void
worksheet::_write_tab_color()
{
    char rgb_str[LXW_ATTR_32];

    if (tab_color == LXW_COLOR_UNSET)
        return;

    lxw_snprintf(rgb_str, LXW_ATTR_32, "FF%06X",
                 tab_color & LXW_COLOR_MASK);

    xml_attribute_list attributes = {
        {"rgb", rgb_str}
    };

    lxw_xml_empty_tag("tabColor", attributes);
}

/*
 * Write the <sheetPr> element for Sheet level properties.
 */
void
worksheet::_write_sheet_pr()
{
    if (!fit_page
        && !filter_on
        && tab_color == LXW_COLOR_UNSET
        && !outline_changed && !vba_codename) {
        return;
    }

    xml_attribute_list attributes;

    if (vba_codename)
        attributes.push_back({"codeName", std::to_string(vba_codename)});

    if (filter_on)
        attributes.push_back({"filterMode", "1"});

    if (fit_page || tab_color != LXW_COLOR_UNSET
        || outline_changed) {
        lxw_xml_start_tag("sheetPr", attributes);
        _write_tab_color();
        /*_write_outline_pr(); */
        _write_page_set_up_pr();
        lxw_xml_end_tag("sheetPr");
    }
    else {
        lxw_xml_empty_tag("sheetPr", attributes);
    }
}

/*
 * Write the <brk> element.
 */
void worksheet::_write_brk(uint32_t id, uint32_t max)
{
    xml_attribute_list attributes = {
        {"id", std::to_string(id)},
        {"max", std::to_string(max)},
        {"man", "1"}
    };

    lxw_xml_empty_tag("brk", attributes);
}

/*
 * Write the <rowBreaks> element.
 */
void worksheet::_write_row_breaks()
{
    uint16_t count = hbreaks_count;
    uint16_t i;

    if (!count)
        return;

    xml_attribute_list attributes = {
        {"count", std::to_string(count)},
        {"manualBreakCount", std::to_string(count)}
    };

    lxw_xml_start_tag("rowBreaks", attributes);

    for (i = 0; i < count; i++)
       _write_brk(hbreaks[i], LXW_COL_MAX - 1);

    lxw_xml_end_tag("rowBreaks");
}

/*
 * Write the <colBreaks> element.
 */
void worksheet::_write_col_breaks()
{
    uint16_t count = vbreaks_count;
    uint16_t i;

    if (!count)
        return;

    xml_attribute_list attributes = {
        {"count", std::to_string(count)},
        {"manualBreakCount", std::to_string(count)}
    };

    lxw_xml_start_tag("colBreaks", attributes);

    for (i = 0; i < count; i++)
       _write_brk(vbreaks[i], LXW_ROW_MAX - 1);

    lxw_xml_end_tag("colBreaks");
}

/*
 * Write the <autoFilter> element.
 */
void worksheet::_write_auto_filter()
{
    std::string range;

    if (!autofilter_.in_use)
        return;

    lxw_rowcol_to_range(range,
                        autofilter_.first_row,
                        autofilter_.first_col,
                        autofilter_.last_row, autofilter_.last_col);

    xml_attribute_list attributes = {
        {"ref", range}
    };

    lxw_xml_empty_tag("autoFilter", attributes);
}

/*
 * Write the <hyperlink> element for external links.
 */
void worksheet::_write_hyperlink_external(lxw_row_t row_num,
                                    lxw_col_t col_num, const std::string& location,
                                    const std::string& tooltip, uint16_t id)
{
    std::string ref;
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_rowcol_to_cell(ref, row_num, col_num);

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", id);

    xml_attribute_list attributes;
    attributes.push_back({"ref", ref});
    attributes.push_back({"r:id", r_id});

    if (!location.empty())
        attributes.push_back({"location", location});

    if (!tooltip.empty())
        attributes.push_back({"tooltip", tooltip});

    lxw_xml_empty_tag("hyperlink", attributes);


}

/*
 * Write the <hyperlink> element for internal links.
 */
void worksheet::_write_hyperlink_internal(lxw_row_t row_num,
                                    lxw_col_t col_num, const std::string& location,
                                    const std::string& display, const std::string& tooltip)
{
    std::string ref;

    lxw_rowcol_to_cell(ref, row_num, col_num);

    xml_attribute_list attributes = {
        {"ref", ref}
    };

    if (!location.empty())
        attributes.push_back({"location", location});

    if (!tooltip.empty())
        attributes.push_back({"tooltip", tooltip});

    if (!display.empty())
        attributes.push_back({"display", display});

    lxw_xml_empty_tag("hyperlink", attributes);
}

/*
 * Process any stored hyperlinks in row/col order and write the <hyperlinks>
 * element. The attributes are different for internal and external links.
 */
void worksheet::_write_hyperlinks()
{

    lxw_row *row;
    lxw_cell *link;
    rel_tuple_ptr relationship;

    if (RB_EMPTY(hyperlinks))
        return;

    /* Write the hyperlink elements. */
    lxw_xml_start_tag("hyperlinks");

    RB_FOREACH(row, lxw_table_rows, hyperlinks) {

        RB_FOREACH(link, lxw_table_cells, row->cells) {

            if (link->type == HYPERLINK_URL
                || link->type == HYPERLINK_EXTERNAL) {

                rel_count++;

                relationship = std::make_shared<rel_tuple>();

                relationship->type = "/hyperlink";

                relationship->target = *link->u.string;

                relationship->target_mode = "External";

                external_hyperlinks.push_back(relationship);

               _write_hyperlink_external(link->row_num,
                                                    link->col_num,
                                                    *link->user_data1,
                                                    *link->user_data2,
                                                    rel_count);
            }

            if (link->type == HYPERLINK_INTERNAL) {

               _write_hyperlink_internal(link->row_num,
                                                    link->col_num,
                                                    *link->u.string,
                                                    *link->user_data1,
                                                    *link->user_data2);
            }

        }

    }
mem_error:
    lxw_xml_end_tag("hyperlinks");
}

/*
 * Write the <sheetProtection> element.
 */
void
worksheet::_write_sheet_protection()
{
    struct lxw_protection *protect = &protection;

    if (!protect->is_configured)
        return;

    xml_attribute_list attributes;

    if (*protect->hash)
        attributes.push_back({"password", protect->hash});

    if (!protect->no_sheet)
        attributes.push_back({"sheet", "1"});

    if (protect->content)
        attributes.push_back({"content", "1"});

    if (!protect->objects)
        attributes.push_back({"objects", "1"});

    if (!protect->scenarios)
        attributes.push_back({"scenarios", "1"});

    if (protect->format_cells)
        attributes.push_back({"formatCells", "0"});

    if (protect->format_columns)
        attributes.push_back({"formatColumns", "0"});

    if (protect->format_rows)
        attributes.push_back({"formatRows", "0"});

    if (protect->insert_columns)
        attributes.push_back({"insertColumns", "0"});

    if (protect->insert_rows)
        attributes.push_back({"insertRows", "0"});

    if (protect->insert_hyperlinks)
        attributes.push_back({"insertHyperlinks", "0"});

    if (protect->delete_columns)
        attributes.push_back({"deleteColumns", "0"});

    if (protect->delete_rows)
        attributes.push_back({"deleteRows", "0"});

    if (protect->no_select_locked_cells)
        attributes.push_back({"selectLockedCells", "1"});

    if (protect->sort)
        attributes.push_back({"sort", "0"});

    if (protect->autofilter)
        attributes.push_back({"autoFilter", "0"});

    if (protect->pivot_tables)
        attributes.push_back({"pivotTables", "0"});

    if (protect->no_select_unlocked_cells)
        attributes.push_back({"selectUnlockedCells", "1"});

    lxw_xml_empty_tag("sheetProtection", attributes);
}

/*
 * Write the <drawing> element.
 */
void worksheet::_write_drawing(uint16_t id)
{
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", id);

    xml_attribute_list attributes = {
        {"r:id", r_id}
    };

    lxw_xml_empty_tag("drawing", attributes);
}

/*
 * Write the <drawing> elements.
 */
void worksheet::_write_drawings()
{
    if (!drawing)
        return;

    rel_count++;

    _write_drawing(rel_count);
}

/*
 * Assemble and write the XML file.
 */
void worksheet::assemble_xml_file()
{
    /* Write the XML declaration. */
    _xml_declaration();

    /* Write the worksheet element. */
    _write_worksheet();

    /* Write the worksheet properties. */
   _write_sheet_pr();

    /* Write the worksheet dimensions. */
   _write_dimension();

    /* Write the sheet view properties. */
   _write_sheet_views();

    /* Write the sheet format properties. */
   _write_sheet_format_pr();

    /* Write the sheet column info. */
   _write_cols();

    /* Write the sheetData element. */
    if (!optimize)
       _write_sheet_data();
    else
       _write_optimized_sheet_data();

    /* Write the sheetProtection element. */
   _write_sheet_protection();

    /* Write the autoFilter element. */
   _write_auto_filter();

    /* Write the mergeCells element. */
   _write_merge_cells();

    /* Write the hyperlink element. */
   _write_hyperlinks();

    /* Write the printOptions element. */
   _write_print_options();

    /* Write the worksheet page_margins. */
   _write_page_margins();

    /* Write the worksheet page setup. */
   _write_page_setup();

    /* Write the headerFooter element. */
   _write_header_footer();

    /* Write the rowBreaks element. */
   _write_row_breaks();

    /* Write the colBreaks element. */
   _write_col_breaks();

    /* Write the drawing element. */
    _write_drawings();

    /* Close the worksheet tag. */
    lxw_xml_end_tag("worksheet");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Write a number to a cell in Excel.
 */
lxw_error
worksheet::write_number(lxw_row_t row_num,
                       lxw_col_t col_num, double value, const format_ptr& format)
{
    lxw_cell *cell;
    lxw_error err;

    err = _check_dimensions(row_num, col_num, false, false);
    if (err)
        return err;

    cell = _new_number_cell(row_num, col_num, value, format.get());

    _insert_cell(row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a string to an Excel file.
 */
lxw_error
worksheet::write_string(lxw_row_t row_num,
                       lxw_col_t col_num, const std::string& string,
                      const format_ptr& format)
{
    lxw_cell *cell;
    int32_t string_id;
    std::string *string_copy = new std::string();
    struct sst_element *sst_element;
    lxw_error err;

    if (string.empty()) {
        /* Treat a NULL or empty string with formatting as a blank cell. */
        /* Null strings without formats should be ignored.      */
        if (format)
            return write_blank(row_num, col_num, format);
        else
            return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    err = _check_dimensions(row_num, col_num, false, false);
    if (err)
        return err;

    if (string.size() > LXW_STR_MAX)
        return LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED;

    if (!optimize) {
        /* Get the SST element and string id. */
        sst_element = sst->get_sst_index(string);

        if (!sst_element)
            return LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND;

        string_id = sst_element->index;
        cell = _new_string_cell(row_num, col_num, string_id, &sst_element->string, format.get());
    }
    else {
        /* Look for and escape control chars in the string. */
        if (strpbrk(string.c_str(), "\x01\x02\x03\x04\x05\x06\x07\x08\x0B\x0C"
                    "\x0D\x0E\x0F\x10\x11\x12\x13\x14\x15\x16"
                    "\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F")) {
            *string_copy = lxw_escape_control_characters(string);
        }
        else {
            *string_copy = string;
        }
        cell = _new_inline_string_cell(row_num, col_num, string_copy, format.get());
    }

    _insert_cell(row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a formula with a numerical result to a cell in Excel.
 */
lxw_error worksheet::write_formula_num(
        lxw_row_t row_num,
        lxw_col_t col_num,
        const std::string& formula,
        const format_ptr& format, double result)
{
    lxw_cell *cell;
    std::string* formula_copy = new std::string();
    lxw_error err;

    if (formula.empty())
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    err = _check_dimensions(row_num, col_num, false, false);
    if (err)
        return err;

    /* Strip leading "=" from formula. */
    if (formula[0] == '=')
        *formula_copy = formula.substr(1);
    else
        *formula_copy = formula;

    cell = _new_formula_cell(row_num, col_num, formula_copy, format.get());
    cell->formula_result = result;

    _insert_cell(row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a formula with a default result to a cell in Excel .
 */
lxw_error
worksheet::write_formula(
        lxw_row_t row_num,
        lxw_col_t col_num,
        const std::string& formula,
        const format_ptr& format)
{
    return write_formula_num(row_num, col_num, formula, format, 0);
}

/*
 * Write a formula with a numerical result to a cell in Excel.
 */
lxw_error worksheet::write_array_formula_num(
        lxw_row_t first_row,
        lxw_col_t first_col,
        lxw_row_t last_row,
        lxw_col_t last_col,
        const std::string& formula,
        const format_ptr& format, double result)
{
    lxw_cell *cell;
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    std::string* formula_copy = new std::string();
    std::string* range = new std::string();
    lxw_error err;

    /* Swap last row/col with first row/col as necessary */
    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }
    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    if (formula.empty())
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    /* Check that column number is valid and store the max value */
    err = _check_dimensions(last_row, last_col, false, false);
    if (err)
        return err;

    /* Define the array range. */

    if (first_row == last_row && first_col == last_col)
        lxw_rowcol_to_cell(*range, first_row, last_col);
    else
        lxw_rowcol_to_range(*range, first_row, first_col, last_row, last_col);

    /* Copy and trip leading "{=" from formula. */
    if (formula[0] == '{')
        if (formula[1] == '=')
            *formula_copy = formula.substr(2);
        else
            *formula_copy = formula.substr(1);
    else
        *formula_copy = formula;

    /* Strip trailing "}" from formula. */
    if ((*formula_copy)[formula_copy->size() - 1] == '}')
        (*formula_copy)[formula_copy->size() - 1] = '\0';

    /* Create a new array formula cell object. */
    cell = _new_array_formula_cell(first_row, first_col, formula_copy, range, format.get());

    cell->formula_result = result;

    _insert_cell(first_row, first_col, cell);

    /* Pad out the rest of the area with formatted zeroes. */
    if (!optimize) {
        for (tmp_row = first_row; tmp_row <= last_row; tmp_row++) {
            for (tmp_col = first_col; tmp_col <= last_col; tmp_col++) {
                if (tmp_row == first_row && tmp_col == first_col)
                    continue;

                write_number(tmp_row, tmp_col, 0, format);
            }
        }
    }

    return LXW_NO_ERROR;
}

/*
 * Write an array formula with a default result to a cell in Excel .
 */
lxw_error worksheet::write_array_formula(
                              lxw_row_t first_row,
                              lxw_col_t first_col,
                              lxw_row_t last_row,
                              lxw_col_t last_col,
                              const std::string& formula, const format_ptr& format)
{
    return write_array_formula_num(first_row, first_col, last_row, last_col, formula, format, 0);
}

/*
 * Write a blank cell with a format to a cell in Excel.
 */
lxw_error worksheet::write_blank(lxw_row_t row_num, lxw_col_t col_num,
                      const format_ptr& format)
{
    lxw_cell *cell;
    lxw_error err;

    /* Blank cells without formatting are ignored by Excel. */
    if (!format)
        return LXW_NO_ERROR;

    err = _check_dimensions(row_num, col_num, false, false);
    if (err)
        return err;

    cell = _new_blank_cell(row_num, col_num, format.get());

    _insert_cell(row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a boolean cell with a format to a cell in Excel.
 */
lxw_error worksheet::write_boolean(lxw_row_t row_num, lxw_col_t col_num, bool value, const format_ptr& format)
{
    lxw_cell *cell;
    lxw_error err;

    err = _check_dimensions(row_num, col_num, false, false);

    if (err)
        return err;

    cell = _new_boolean_cell(row_num, col_num, value, format.get());

    _insert_cell(row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a date and or time to a cell in Excel.
 */
lxw_error worksheet::write_datetime(
                         lxw_row_t row_num,
                         lxw_col_t col_num, lxw_datetime *datetime,
                         const format_ptr& format)
{
    lxw_cell *cell;
    double excel_date;
    lxw_error err;

    err = _check_dimensions(row_num, col_num, false, false);
    if (err)
        return err;

    excel_date = lxw_datetime_to_excel_date(datetime, LXW_EPOCH_1900);

    cell = _new_number_cell(row_num, col_num, excel_date, format.get());

    _insert_cell(row_num, col_num, cell);

    return LXW_NO_ERROR;
}

/*
 * Write a hyperlink/url to an Excel file.
 */
lxw_error
worksheet::write_url_opt(lxw_row_t row_num,
                        lxw_col_t col_num, const std::string& url,
                        const format_ptr& format, const std::string& string,
                        const std::string& tooltip)
{
    lxw_cell *link;
    std::string *string_copy = new std::string();
    std::string *url_copy = nullptr;
    std::string *url_external = nullptr;
    std::string *url_string = nullptr;
    std::string *tooltip_copy = nullptr;
    std::string *found_string = nullptr;
    lxw_error err;
    size_t string_size;
    enum cell_types link_type = HYPERLINK_URL;

    if (url.empty())
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    /* Check the Excel limit of URLS per worksheet. */
    if (hlink_count > LXW_MAX_NUMBER_URLS)
        return LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED;

    err = _check_dimensions(row_num, col_num, false, false);
    if (err)
        return err;

    /* Set the URI scheme from internal links. */
    size_t idx = url.find("internal:");
    if (idx < url.size())
        link_type = HYPERLINK_INTERNAL;

    /* Set the URI scheme from external links. */
    idx = url.find("external:");
    if (idx < url.size())
        link_type = HYPERLINK_EXTERNAL;

    if (!string.empty()) {
        *string_copy = string;
    }
    else {
        if (link_type == HYPERLINK_URL) {
            /* Strip the mailto header. */
            idx = url.find("mailto:");
            if (idx < url.size())
                *string_copy = url.substr(sizeof("mailto"));
            else
                *string_copy = url;
        }
        else {
            *string_copy = url.substr(sizeof("__ternal"));
        }
    }

    if (!url.empty()) {
        url_copy = new std::string();
        if (link_type == HYPERLINK_URL)
            *url_copy = url;
        else
            *url_copy = url.substr(sizeof("__ternal"));
    }

    if (!tooltip.empty()) {
        tooltip_copy = new std::string();
        *tooltip_copy = tooltip;
    }

    if (link_type == HYPERLINK_INTERNAL) {
        url_string = new std::string();
        *url_string = *string_copy;
    }

    /* Escape the URL. */
    if (link_type == HYPERLINK_URL && url_copy->size() >= 3) {
        uint8_t not_escaped = 1;

        /* First check if the URL is already escaped by the user. */
        for (size_t i = 0; i <= url_copy->size() - 3; i++) {
            if ((*url_copy)[i] == '%' && isxdigit((*url_copy)[i + 1])
                && isxdigit((*url_copy)[i + 2])) {

                not_escaped = 0;
                break;
            }
        }

        if (not_escaped) {
            url_external = new std::string();

            for (size_t i = 0; i <= url_copy->size(); i++) {
                switch ((*url_copy)[i]) {
                case (' '):
                case ('"'):
                case ('%'):
                case ('<'):
                case ('>'):
                case ('['):
                case (']'):
                case ('`'):
                case ('^'):
                case ('{'):
                case ('}'):
                {
                    std::stringstream ss;
                    ss << "%" << std::hex << (int)(*url_copy)[i];
                    url_external->append(ss.str());
                    break;
                }
                default:
                    url_external->push_back((*url_copy)[i]);
                }

            }

            delete url_copy;
            url_copy = url_external;

            url_external = nullptr;
        }
    }

    if (link_type == HYPERLINK_EXTERNAL) {
        /* External Workbook links need to be modified into the right format.
         * The URL will look something like "c:\temp\file.xlsx#Sheet!A1".
         * We need the part to the left of the # as the URL and the part to
         * the right as the "location" string (if it exists).
         */

        /* For external links change the dir separator from Unix to DOS. */
        for (size_t i = 0; i < url_copy->size(); i++)
            if ((*url_copy)[i] == '/')
                (*url_copy)[i] = '\\';

        for (size_t i = 0; i < string_copy->size(); i++)
            if ((*string_copy)[i] == '/')
                (*string_copy)[i] = '\\';

        idx = url_copy->find('#');

        if (idx < url_copy->size()) {
            *url_string = url_copy->substr(idx + 1);

            *url_copy = url_copy->substr(0, idx);
        }

        /* Look for Windows style "C:/" link or Windows share "\\" link. */
        idx  = url_copy->find(':');
        if (idx == url_copy->size())
            idx = url_copy->find("\\\\");

        if (idx < url_copy->size()) {
            /* Add the file:/// URI to the url if non-local. */
            string_size = sizeof("file:///") + url_copy->size();
            url_external = new std::string();

            *url_external = "file:///" + *url_copy;

        }

        /* Convert a ./dir/file.xlsx link to dir/file.xlsx. */
        idx = url_copy->find(".\\");
        if (idx < url_copy->size())
            *url_copy = url_copy->substr(2);

        if (url_external) {
            url_copy = url_external;

            delete url_external;
            url_external = NULL;
        }

    }

    /* Excel limits escaped URL to 255 characters. */
    if (url_copy->size() > 255)
        //! @TODO make log here
        delete string_copy;
        return LXW_NO_ERROR;

    err = write_string(row_num, col_num, *string_copy, format);
    if (err)
        //! @TODO make log here
        delete string_copy;
        return LXW_NO_ERROR;

    link = _new_hyperlink_cell(row_num, col_num, link_type, url_copy,
                               url_string, tooltip_copy);

    _insert_hyperlink(row_num, col_num, link);

    delete string_copy;
    hlink_count++;
    return LXW_NO_ERROR;
}

/*
 * Write a hyperlink/url to an Excel file.
 */
lxw_error worksheet::write_url(lxw_row_t row_num,
                     lxw_col_t col_num,
                     const std::string& url,
                     const format_ptr& format)
{
    return write_url_opt(row_num, col_num, url, format);
}

/*
 * Set the properties of a single column or a range of columns with options.
 */
lxw_error worksheet::set_column_opt(
                         lxw_col_t firstcol,
                         lxw_col_t lastcol,
                         double width,
                         const format_ptr& format,
                         const lxw_row_col_options& user_options)
{
    lxw_col_options *copied_options;
    uint8_t ignore_row = true;
    uint8_t ignore_col = true;
    bool hidden = user_options.hidden;
    uint8_t level = user_options.level;
    bool collapsed = user_options.collapsed;
    lxw_col_t col;
    lxw_error err;

    /* Ensure second col is larger than first. */
    if (firstcol > lastcol) {
        lxw_col_t tmp = firstcol;
        firstcol = lastcol;
        lastcol = tmp;
    }

    /* Ensure that the cols are valid and store max and min values.
     * NOTE: The check shouldn't modify the row dimensions and should only
     *       modify the column dimensions in certain cases. */
    if (format != NULL || (width != LXW_DEF_COL_WIDTH && hidden))
        ignore_col = false;

    err = _check_dimensions(0, firstcol, ignore_row, ignore_col);

    if (!err)
        err = _check_dimensions(0, lastcol, ignore_row, ignore_col);

    if (err)
        return err;

    /* Resize the col_options array if required. */
    if (firstcol >= col_options_max) {
        lxw_col_t col;
        lxw_col_t old_size = col_options_max;
        lxw_col_t new_size = _next_power_of_two(firstcol + 1);
        lxw_col_options **new_ptr = (lxw_col_options **)realloc(col_options,
                                            new_size *
                                            sizeof(lxw_col_options *));

        if (new_ptr) {
            for (col = old_size; col < new_size; col++)
                new_ptr[col] = NULL;

            col_options = new_ptr;
            col_options_max = new_size;
        }
        else {
            return LXW_ERROR_MEMORY_MALLOC_FAILED;
        }
    }

    /* Resize the col_formats array if required. */
    if (lastcol >= col_formats_max) {
        lxw_col_t col;
        lxw_col_t old_size = col_formats_max;
        lxw_col_t new_size = _next_power_of_two(lastcol + 1);
        xlsxwriter::format **new_ptr = (xlsxwriter::format **)realloc(col_formats, new_size * sizeof(xlsxwriter::format*));

        if (new_ptr) {
            for (col = old_size; col < new_size; col++)
                new_ptr[col] = NULL;

            col_formats = new_ptr;
            col_formats_max = new_size;
        }
        else {
            return LXW_ERROR_MEMORY_MALLOC_FAILED;
        }
    }

    /* Store the column options. */
    copied_options = new lxw_col_options();

    copied_options->firstcol = firstcol;
    copied_options->lastcol = lastcol;
    copied_options->width = width;
    copied_options->format = format;
    copied_options->hidden = hidden;
    copied_options->level = level;
    copied_options->collapsed = collapsed;

    col_options[firstcol] = copied_options;

    /* Store the column formats for use when writing cell data. */
    for (col = firstcol; col <= lastcol; col++) {
        col_formats[col] = format.get();
    }

    /* Store the column change to allow optimizations. */
    col_size_changed = true;

    return LXW_NO_ERROR;
}

/*
 * Set the properties of a single column or a range of columns.
 */
lxw_error worksheet::set_column(lxw_col_t firstcol, lxw_col_t lastcol, double width, const format_ptr& format)
{
    return set_column_opt(firstcol, lastcol, width, format);
}

/*
 * Set the properties of a row with options.
 */
lxw_error worksheet::set_row_opt( lxw_row_t row_num, double height, const format_ptr& format, const lxw_row_col_options& user_options)
{

    lxw_col_t min_col;
    bool hidden = user_options.hidden;
    uint8_t level = user_options.level;
    bool collapsed = user_options.collapsed;
    lxw_row *row;
    lxw_error err;

    /* Use minimum col in _check_dimensions(). */
    if (dim_colmin != LXW_COL_MAX)
        min_col = dim_colmin;
    else
        min_col = 0;

    err = _check_dimensions(row_num, min_col, false, false);
    if (err)
        return err;

    /* If the height is 0 the row is hidden and the height is the default. */
    if (height == 0) {
        hidden = true;
        height = default_row_height;
    }

    row = _get_row(row_num);

    row->height = height;
    row->format = format.get();
    row->hidden = hidden;
    row->level = level;
    row->collapsed = collapsed;
    row->row_changed = true;

    if (height != default_row_height)
        row->height_changed = true;

    return LXW_NO_ERROR;
}

/*
 * Set the properties of a row.
 */
lxw_error worksheet::set_row(lxw_row_t row_num, double height, const format_ptr& format)
{
    return set_row_opt(row_num, height, format);
}

/*
 * Merge a range of cells. The first cell should contain the data and the others
 * should be blank. All cells should contain the same format.
 */
lxw_error worksheet::merge_range(lxw_row_t first_row,
                      lxw_col_t first_col, lxw_row_t last_row,
                      lxw_col_t last_col, const std::string& string,
                      const format_ptr& format)
{
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    lxw_error err;

    /* Excel doesn't allow a single cell to be merged */
    if (first_row == last_row && first_col == last_col)
        return LXW_ERROR_PARAMETER_VALIDATION;

    /* Swap last row/col with first row/col as necessary */
    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }
    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    /* Check that column number is valid and store the max value */
    err = _check_dimensions(last_row, last_col, false, false);
    if (err)
        return err;

    /* Store the merge range. */
    std::shared_ptr<lxw_merged_range> merged_range = std::make_shared<lxw_merged_range>();

    merged_range->first_row = first_row;
    merged_range->first_col = first_col;
    merged_range->last_row = last_row;
    merged_range->last_col = last_col;

    merged_ranges.push_back(merged_range);
    merged_range_count++;

    /* Write the first cell */
    write_string(first_row, first_col, string, format);

    /* Pad out the rest of the area with formatted blank cells. */
    for (tmp_row = first_row; tmp_row <= last_row; tmp_row++) {
        for (tmp_col = first_col; tmp_col <= last_col; tmp_col++) {
            if (tmp_row == first_row && tmp_col == first_col)
                continue;
            write_blank(tmp_row, tmp_col, format);
        }
    }

    return LXW_NO_ERROR;
}

/*
 * Set the autofilter area in the worksheet.
 */
lxw_error worksheet::autofilter(lxw_row_t first_row,
                     lxw_col_t first_col, lxw_row_t last_row,
                     lxw_col_t last_col)
{
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    lxw_error err;

    /* Excel doesn't allow a single cell to be merged */
    if (first_row == last_row && first_col == last_col)
        return LXW_ERROR_PARAMETER_VALIDATION;

    /* Swap last row/col with first row/col as necessary */
    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }
    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    /* Check that column number is valid and store the max value */
    err = _check_dimensions(last_row, last_col, false, false);
    if (err)
        return err;

    autofilter_.in_use = true;
    autofilter_.first_row = first_row;
    autofilter_.first_col = first_col;
    autofilter_.last_row = last_row;
    autofilter_.last_col = last_col;

    return LXW_NO_ERROR;
}

/*
 * Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
 * highlighted.
 */
void
worksheet::select()
{
    selected = true;

    /* Selected worksheet can't be hidden. */
    hidden = false;
}

/*
 * Set this worksheet as the active worksheet, i.e. the worksheet that is
 * displayed when the workbook is opened. Also set it as selected.
 */
void
worksheet::activate()
{
    selected = true;
    active = true;

    /* Active worksheet can't be hidden. */
    hidden = false;

    *active_sheet = index;
}

/*
 * Set this worksheet as the first visible sheet. This is necessary
 * when there are a large number of worksheets and the activated
 * worksheet is not visible on the screen.
 */
void
worksheet::set_first_sheet()
{
    /* Active worksheet can't be hidden. */
    hidden = false;

    *first_sheet = index;
}

/*
 * Hide this worksheet.
 */
void
worksheet::hide()
{
    hidden = true;

    /* A hidden worksheet shouldn't be active or selected. */
    selected = false;

    /* If this is active_sheet or first_sheet reset the workbook value. */
    if (*first_sheet == index)
        *first_sheet = 0;

    if (*active_sheet == index)
        *active_sheet = 0;
}

/*
 * Set which cell or cells are selected in a worksheet.
 */
void worksheet::set_selection(lxw_row_t first_row, lxw_col_t first_col,
                         lxw_row_t last_row, lxw_col_t last_col)
{
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    std::string active_cell("", LXW_MAX_CELL_RANGE_LENGTH);
    std::string sqref("", LXW_MAX_CELL_RANGE_LENGTH);

    /* Only allow selection to be set once to avoid freeing/re-creating it. */
    if (!selections.empty())
        return;

    /* Excel doesn't set a selection for cell A1 since it is the default. */
    if (first_row == 0 && first_col == 0 && last_row == 0 && last_col == 0)
        return;

    std::shared_ptr<lxw_selection> selection = std::make_shared<lxw_selection>();

    /* Set the cell range selection. Do this before swapping max/min to  */
    /* allow the selection direction to be reversed. */
    lxw_rowcol_to_cell(active_cell, first_row, first_col);

    /* Swap last row/col for first row/col if necessary. */
    if (first_row > last_row) {
        tmp_row = first_row;
        first_row = last_row;
        last_row = tmp_row;
    }

    if (first_col > last_col) {
        tmp_col = first_col;
        first_col = last_col;
        last_col = tmp_col;
    }

    /* If the first and last cell are the same write a single cell. */
    if ((first_row == last_row) && (first_col == last_col))
        lxw_rowcol_to_cell(sqref, first_row, first_col);
    else
        lxw_rowcol_to_range(sqref, first_row, first_col, last_row, last_col);

    selection->pane = "";
    selection->active_cell = active_cell;
    selection->sqref = sqref;

    selections.push_back(selection);
}

/*
 * Set panes and mark them as frozen. With extra options.
 */
void worksheet::freeze_panes_opt( lxw_row_t first_row, lxw_col_t first_col,
                           lxw_row_t top_row, lxw_col_t left_col,
                           uint8_t type)
{
    panes.first_row = first_row;
    panes.first_col = first_col;
    panes.top_row = top_row;
    panes.left_col = left_col;
    panes.x_split = 0.0;
    panes.y_split = 0.0;

    if (type)
        panes.type = FREEZE_SPLIT_PANES;
    else
        panes.type = FREEZE_PANES;
}

/*
 * Set panes and mark them as frozen.
 */
void worksheet::freeze_panes(lxw_row_t first_row, lxw_col_t first_col)
{
    freeze_panes_opt(first_row, first_col, first_row, first_col, 0);
}

/*
 * Set panes and mark them as split.With extra options.
 */
void worksheet::split_panes_opt(double y_split, double x_split,
                          lxw_row_t top_row, lxw_col_t left_col)
{
    panes.first_row = 0;
    panes.first_col = 0;
    panes.top_row = top_row;
    panes.left_col = left_col;
    panes.x_split = x_split;
    panes.y_split = y_split;
    panes.type = SPLIT_PANES;
}

/*
 * Set panes and mark them as split.
 */
void worksheet::split_panes(double y_split, double x_split)
{
    split_panes_opt(y_split, x_split, 0, 0);
}

/*
 * Set the page orientation as portrait.
 */
void worksheet::set_portrait()
{
    orientation = LXW_PORTRAIT;
    page_setup_changed = true;
}

/*
 * Set the page orientation as landscape.
 */
void worksheet::set_landscape()
{
    orientation = LXW_LANDSCAPE;
    page_setup_changed = true;
}

/*
 * Set the page view mode for Mac Excel.
 */
void
worksheet::set_page_view()
{
    page_view = true;
}

/*
 * Set the paper type. Example. 1 = US Letter, 9 = A4
 */
void worksheet::set_paper(uint8_t paper_size)
{
    paper_size = paper_size;
    page_setup_changed = true;
}

/*
 * Set the order in which pages are printed.
 */
void worksheet::print_across()
{
    page_order = LXW_PRINT_ACROSS;
    page_setup_changed = true;
}

/*
 * Set all the page margins in inches.
 */
void worksheet::set_margins(double left, double right,
                      double top, double bottom)
{

    if (left >= 0)
        margin_left = left;

    if (right >= 0)
        margin_right = right;

    if (top >= 0)
        margin_top = top;

    if (bottom >= 0)
        margin_bottom = bottom;
}

/*
 * Set the page header caption and options.
 */
lxw_error worksheet::set_header_opt(const std::string& string,
                         const lxw_header_footer_options& options)
{
    if (options.margin > 0)
        margin_header = options.margin;


    if (string.empty())
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    if (string.size() >= LXW_HEADER_FOOTER_MAX)
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;

    header = string;
    header_footer_changed = 1;

    return LXW_NO_ERROR;
}

/*
 * Set the page footer caption and options.
 */
lxw_error
worksheet::set_footer_opt(const std::string& string, const lxw_header_footer_options& options)
{
    if (options.margin > 0)
        margin_footer = options.margin;


    if (string.empty())
        return LXW_ERROR_NULL_PARAMETER_IGNORED;

    if (string.size() >= LXW_HEADER_FOOTER_MAX)
        return LXW_ERROR_255_STRING_LENGTH_EXCEEDED;

    footer = string;
    header_footer_changed = 1;

    return LXW_NO_ERROR;
}

/*
 * Set the page header caption.
 */
lxw_error
worksheet::set_header(const std::string& string)
{
    return set_header_opt(string);
}

/*
 * Set the page footer caption.
 */
lxw_error
worksheet::set_footer(const std::string& string)
{
    return set_footer_opt(string);
}

/*
 * Set the option to show/hide gridlines on the screen and the printed page.
 */
void
worksheet::gridlines(uint8_t option)
{
    if (option == LXW_HIDE_ALL_GRIDLINES) {
        print_gridlines = 0;
        screen_gridlines = 0;
    }

    if (option & LXW_SHOW_SCREEN_GRIDLINES) {
        screen_gridlines = 1;
    }

    if (option & LXW_SHOW_PRINT_GRIDLINES) {
        print_gridlines = 1;
        print_options_changed = 1;
    }
}

/*
 * Center the page horizontally.
 */
void
worksheet::center_horizontally()
{
    print_options_changed = 1;
    hcenter = 1;
}

/*
 * Center the page horizontally.
 */
void
worksheet::center_vertically()
{
    print_options_changed = 1;
    vcenter = 1;
}

/*
 * Set the option to print the row and column headers on the printed page.
 */
void
worksheet::print_row_col_headers()
{
    print_headers = 1;
    print_options_changed = 1;
}

/*
 * Set the rows to repeat at the top of each printed page.
 */
lxw_error
worksheet::repeat_rows(lxw_row_t first_row,
                      lxw_row_t last_row)
{
    lxw_row_t tmp_row;
    lxw_error err;

    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }

    err = _check_dimensions(last_row, 0, LXW_IGNORE, LXW_IGNORE);
    if (err)
        return err;

    repeat_rows_.in_use = true;
    repeat_rows_.first_row = first_row;
    repeat_rows_.last_row = last_row;

    return LXW_NO_ERROR;
}

/*
 * Set the columns to repeat at the left hand side of each printed page.
 */
lxw_error
worksheet::repeat_columns(lxw_col_t first_col,
                         lxw_col_t last_col)
{
    lxw_col_t tmp_col;
    lxw_error err;

    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    err = _check_dimensions(last_col, 0, LXW_IGNORE, LXW_IGNORE);
    if (err)
        return err;

    repeat_cols_.in_use = true;
    repeat_cols_.first_col = first_col;
    repeat_cols_.last_col = last_col;

    return LXW_NO_ERROR;
}

/*
 * Set the print area in the current worksheet.
 */
lxw_error
worksheet::print_area(lxw_row_t first_row,
                     lxw_col_t first_col, lxw_row_t last_row,
                     lxw_col_t last_col)
{
    lxw_row_t tmp_row;
    lxw_col_t tmp_col;
    lxw_error err;

    if (first_row > last_row) {
        tmp_row = last_row;
        last_row = first_row;
        first_row = tmp_row;
    }

    if (first_col > last_col) {
        tmp_col = last_col;
        last_col = first_col;
        first_col = tmp_col;
    }

    err = _check_dimensions(last_row, last_col, LXW_IGNORE, LXW_IGNORE);
    if (err)
        return err;

    /* Ignore max area since it is the same as no print area in Excel. */
    if (first_row == 0 && first_col == 0 && last_row == LXW_ROW_MAX - 1
        && last_col == LXW_COL_MAX - 1) {
        return LXW_NO_ERROR;
    }

    print_area_.in_use = true;
    print_area_.first_row = first_row;
    print_area_.last_row = last_row;
    print_area_.first_col = first_col;
    print_area_.last_col = last_col;

    return LXW_NO_ERROR;
}

/* Store the vertical and horizontal number of pages that will define the
 * maximum area printed.
 */
void
worksheet::fit_to_pages(uint16_t width, uint16_t height)
{
    fit_page = 1;
    fit_width = width;
    fit_height = height;
    page_setup_changed = 1;
}

/*
 * Set the start page number.
 */
void
worksheet::set_start_page(uint16_t start_page)
{
    page_start = start_page;
}

/*
 * Set the scale factor for the printed page.
 */
void
worksheet::set_print_scale(uint16_t scale)
{
    /* Confine the scale to Excel"s range */
    if (scale < 10 || scale > 400)
        return;

    /* Turn off "fit to page" option. */
    fit_page = false;

    print_scale = scale;
    page_setup_changed = true;
}

/*
 * Store the horizontal page breaks on a worksheet.
 */
lxw_error worksheet::set_h_pagebreaks(const std::vector<lxw_row_t>& hbreaks)
{
    uint16_t count = hbreaks.size();

    /* The Excel 2007 specification says that the maximum number of page
     * breaks is 1026. However, in practice it is actually 1023. */
    if (count > LXW_BREAKS_MAX)
        count = LXW_BREAKS_MAX;

    this->hbreaks = hbreaks;
    hbreaks_count = count;

    return LXW_NO_ERROR;
}

/*
 * Store the vertical page breaks on a worksheet.
 */
lxw_error worksheet::set_v_pagebreaks(const std::vector<lxw_col_t>& vbreaks)
{
    uint16_t count = vbreaks.size();

    /* The Excel 2007 specification says that the maximum number of page
     * breaks is 1026. However, in practice it is actually 1023. */
    if (count > LXW_BREAKS_MAX)
        count = LXW_BREAKS_MAX;

    this->vbreaks = vbreaks;
    vbreaks_count = count;

    return LXW_NO_ERROR;
}

/*
 * Set the worksheet zoom factor.
 */
void
worksheet::set_zoom(uint16_t scale)
{
    /* Confine the scale to Excel"s range */
    if (scale < 10 || scale > 400) {
        LXW_WARN("worksheet_set_zoom(): "
                 "Zoom factor scale outside range: 10 <= zoom <= 400.");
        return;
    }

    zoom = scale;
}

/*
 * Hide cell zero values.
 */
void worksheet::hide_zero()
{
    show_zeros = false;
}

/*
 * Display the worksheet right to left for some eastern versions of Excel.
 */
void worksheet::get_right_to_left()
{
    right_to_left = true;
}

/*
 * Set the color of the worksheet tab.
 */
void
worksheet::set_tab_color(lxw_color_t color)
{
    tab_color = color;
}

/*
 * Set the worksheet protection flags to prevent modification of worksheet
 * objects.
 */
void
worksheet::protect(const char *password,
                  lxw_protection *options)
{
    struct lxw_protection *protect = &protection;

    /* Copy any user parameters to the internal structure. */
    if (options)
        memcpy(protect, options, sizeof(lxw_protection));

    /* Zero the hash storage in case of copied initialization data. */
    protect->hash[0] = '\0';

    if (password) {
        uint16_t hash = _hash_password(password);
        lxw_snprintf(protect->hash, 5, "%X", hash);
    }

    protect->is_configured = true;
}

/*
 * Set the default row properties
 */
void
worksheet::set_default_row(double height,
                          uint8_t hide_unused_rows)
{
    if (height < 0)
        height = default_row_height;

    if (height != default_row_height) {
        default_row_height = height;
        row_size_changed = true;
    }

    if (hide_unused_rows)
        default_row_zeroed = true;

    default_row_set = true;
}

/*
 * Insert an image into the worksheet.
 */
lxw_error worksheet::insert_image_opt(
                           lxw_row_t row_num, lxw_col_t col_num,
                           const std::string& filename,
                           const image_options_ptr& user_options)
{
    FILE *image_stream;
    std::string short_name;

    if (filename.empty()) {
        LXW_WARN("worksheet_insert_image()/_opt(): "
                 "filename must be specified.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the image file exists and can be opened. */
    image_stream = fopen(filename.c_str(), "rb");
    if (!image_stream) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "file doesn't exist or can't be opened: %s.",
                         filename);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Get the filename from the full path to add to the Drawing object. */
    short_name = lxw_basename(filename);
    if (short_name.empty()) {
        LXW_WARN_FORMAT1("worksheet_insert_image()/_opt(): "
                         "couldn't get basename for file: %s.", filename);
        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Create a new object to hold the image options. */
    image_options_ptr options = std::make_shared<image_options>();

    if (user_options) {
        *options = *user_options;
        options->url = user_options->url;
        options->tip = user_options->tip;
    }

    /* Copy other options or set defaults. */
    options->filename = filename;
    options->short_name = short_name;
    options->stream = image_stream;
    options->row = row_num;
    options->col = col_num;

    if (!options->x_scale)
        options->x_scale = 1;

    if (!options->y_scale)
        options->y_scale = 1;

    if (_get_image_properties(options) == LXW_NO_ERROR) {
        image_data.push_back(options);
        return LXW_NO_ERROR;
    }
    else {
        return LXW_ERROR_IMAGE_DIMENSIONS;
    }
}

/*
 * Insert an image into the worksheet.
 */
lxw_error worksheet::insert_image( lxw_row_t row_num, lxw_col_t col_num, const std::string& filename)
{
    return insert_image_opt(row_num, col_num, filename);
}

/*
 * Insert an chart into the worksheet.
 */
lxw_error worksheet::insert_chart_opt(lxw_row_t row_num, lxw_col_t col_num, const chart_ptr& chart, const image_options_ptr& user_options)
{
    if (!chart) {
        LXW_WARN("worksheet_insert_chart()/_opt(): chart must be non-NULL.");
        return LXW_ERROR_NULL_PARAMETER_IGNORED;
    }

    /* Check that the chart isn't being used more than once. */
    if (chart->in_use) {
        LXW_WARN("worksheet_insert_chart()/_opt(): the same chart object "
                 "cannot be inserted in a worksheet more than once.");

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Check that the chart has a data series. */
    if (chart->series_list.empty()) {
        LXW_WARN
            ("worksheet_insert_chart()/_opt(): chart must have a series.");

        return LXW_ERROR_PARAMETER_VALIDATION;
    }

    /* Check that the chart has a 'values' series. */
    for (const auto& series : chart->series_list) {
        if (series->values->formula.empty() && series->values->sheetname.empty()) {
            LXW_WARN("worksheet_insert_chart()/_opt(): chart must have a "
                     "'values' series.");

            return LXW_ERROR_PARAMETER_VALIDATION;
        }
    }

    /* Create a new object to hold the chart image options. */
    image_options_ptr options = std::make_shared<image_options>();

    if (user_options)
        memcpy(options.get(), user_options.get(), sizeof(image_options));

    /* Copy other options or set defaults. */
    options->row = row_num;
    options->col = col_num;

    /* TODO. Read defaults from chart. */
    options->width = 480;
    options->height = 288;

    if (!options->x_scale)
        options->x_scale = 1;

    if (!options->y_scale)
        options->y_scale = 1;

    /* Store chart references so they can be ordered in the workbook. */
    options->chart = chart;

    chart_data.push_back(options);

    chart->in_use = true;

    return LXW_NO_ERROR;
}

/*
 * Insert an image into the worksheet.
 */
lxw_error worksheet::insert_chart(lxw_row_t row_num, lxw_col_t col_num, const chart_ptr& chart)
{
    return insert_chart_opt(row_num, col_num, chart, NULL);
}

} // xlsxwriter
