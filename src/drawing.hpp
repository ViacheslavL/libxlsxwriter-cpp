/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * drawing - A libxlsxwriter library for creating Excel XLSX drawing files.
 *
 */
#ifndef __LXW_DRAWING_HPP__
#define __LXW_DRAWING_HPP__

#include <stdint.h>

#include "common.hpp"
#include "xmlwriter.hpp"

namespace xlsxwriter {

enum lxw_drawing_types {
    LXW_DRAWING_NONE = 0,
    LXW_DRAWING_IMAGE,
    LXW_DRAWING_CHART,
    LXW_DRAWING_SHAPE
};

enum lxw_anchor_types {
    LXW_ANCHOR_TYPE_NONE = 0,
    LXW_ANCHOR_TYPE_IMAGE,
    LXW_ANCHOR_TYPE_CHART
};

enum lxw_anchor_edit_types {
    LXW_ANCHOR_EDIT_AS_NONE = 0,
    LXW_ANCHOR_EDIT_AS_RELATIVE,
    LXW_ANCHOR_EDIT_AS_ONE_CELL,
    LXW_ANCHOR_EDIT_AS_ABSOLUTE
};

enum image_types {
    LXW_IMAGE_UNKNOWN = 0,
    LXW_IMAGE_PNG,
    LXW_IMAGE_JPEG,
    LXW_IMAGE_BMP
};

/* Coordinates used in a drawing object. */
struct lxw_drawing_coords {
    uint32_t col;
    uint32_t row;
    double col_offset;
    double row_offset;
};

/* Object to represent the properties of a drawing. */
struct drawing_object {
    uint8_t anchor_type;
    uint8_t edit_as;
    lxw_drawing_coords from;
    lxw_drawing_coords to;
    uint32_t col_absolute;
    uint32_t row_absolute;
    uint32_t width;
    uint32_t height;
    uint8_t shape;
    std::string description;
    std::string url;
    std::string tip;
};

typedef std::shared_ptr<drawing_object> drawing_object_ptr;

class packager;
class worksheet;
/*
 * Struct to represent a collection of drawings.
 */
struct drawing : public xmlwriter {
    friend class packager;
    friend class worksheet;
public:
    void assemble_xml_file();
    void add_drawing_object(const drawing_object_ptr& drawing_object);

    /* Declarations required for unit testing. */

    void _xml_declaration();

private:

    uint8_t embedded;
    std::list<drawing_object_ptr> drawing_objects;

    void _write_drawing_workspace();
    void _write_two_cell_anchor(uint16_t index, const drawing_object_ptr &drawing_object);
    void _write_graphic_frame(uint16_t index);
    void _write_a_graphic(uint16_t index);
    void _write_a_graphic_data(uint16_t index);
    void _write_chart(uint16_t index);
    void _write_xfrm();
    void _write_xfrm_extension();
    void _write_nv_graphic_frame_pr(uint16_t index);
    void _write_xfrm_offset();
    void _write_c_nv_graphic_frame_pr();
    void _write_client_data();
    void _write_pic(uint16_t index, const drawing_object_ptr &drawing_object);
    void _write_sp_pr(const drawing_object_ptr &drawing_object);
    void _write_a_prst_geom();
    void _write_a_xfrm(const drawing_object_ptr &drawing_object);
    void _write_a_off(const drawing_object_ptr &drawing_object);
    void _write_a_ext(const drawing_object_ptr &drawing_object);
    void _write_blip_fill(uint16_t index);
    void _write_a_stretch();
    void _write_a_fill_rect();
    void _write_a_blip(uint16_t index);
    void _write_c_nv_pic_pr();
    void _write_nv_pic_pr(uint16_t index, const drawing_object_ptr &drawing_object);
    void _write_a_pic_locks();
    void _write_c_nv_pr(const std::string &object_name, uint16_t index, const drawing_object_ptr &drawing_object);
    void _write_to(lxw_drawing_coords *coords);
    void _write_from(lxw_drawing_coords *coords);
    void _write_row_off(const std::string &data);
    void _write_col(const std::string &data);
    void _write_col_off(const std::string &data);
    void _write_row(const std::string &data);
    void _write_a_av_lst();
};


} // namespace xlsxwriter

#endif /* __LXW_DRAWING_HPP__ */
