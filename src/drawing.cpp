/*****************************************************************************
 * drawing - A library for creating Excel XLSX drawing files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <xlsxwriter/xmlwriter.hpp>
#include <xlsxwriter/common.hpp>
#include <xlsxwriter/drawing.hpp>
#include <xlsxwriter/utility.hpp>

#define LXW_OBJ_NAME_LENGTH 14  /* "Picture 65536", or "Chart 65536" */
/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

namespace xlsxwriter {

/*
 * Add a drawing object to the drawing collection.
 */
void drawing::add_drawing_object(const drawing_object_ptr& drawing_object)
{
    drawing_objects.push_back(drawing_object);
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
void drawing::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <xdr:wsDr> element.
 */
void drawing::_write_drawing_workspace()
{
    char xmlns_xdr[] = LXW_SCHEMA_DRAWING "/spreadsheetDrawing";
    char xmlns_a[] = LXW_SCHEMA_DRAWING "/main";
    xml_attribute_list attributes = {
        {"xmlns:xdr", xmlns_xdr},
        {"xmlns:a", xmlns_a}
    };

    lxw_xml_start_tag("xdr:wsDr", attributes);
}

/*
 * Write the <xdr:col> element.
 */
void drawing::_write_col(const std::string& data)
{
    lxw_xml_data_element("xdr:col", data);
}

/*
 * Write the <xdr:colOff> element.
 */
void drawing::_write_col_off(const std::string& data)
{
    lxw_xml_data_element("xdr:colOff", data);
}

/*
 * Write the <xdr:row> element.
 */
void drawing::_write_row(const std::string& data)
{
    lxw_xml_data_element("xdr:row", data);
}

/*
 * Write the <xdr:rowOff> element.
 */
void drawing::_write_row_off(const std::string& data)
{
    lxw_xml_data_element("xdr:rowOff", data);
}

/*
 * Write the <xdr:from> element.
 */
void drawing::_write_from(lxw_drawing_coords *coords)
{
    lxw_xml_start_tag("xdr:from");

    _write_col(std::to_string(coords->col));

    _write_col_off(to_string(coords->col_offset));

    _write_row(std::to_string(coords->row));

    _write_row_off(to_string(coords->row_offset));

    lxw_xml_end_tag("xdr:from");
}

/*
 * Write the <xdr:to> element.
 */
void drawing::_write_to(lxw_drawing_coords *coords)
{
    lxw_xml_start_tag("xdr:to");
    _write_col(std::to_string(coords->col));

    _write_col_off(to_string(coords->col_offset));

    _write_row(std::to_string(coords->row));

    _write_row_off(to_string(coords->row_offset));

    lxw_xml_end_tag("xdr:to");
}

/*
 * Write the <xdr:cNvPr> element.
 */
void drawing::_write_c_nv_pr(const std::string& object_name, uint16_t index, const drawing_object_ptr& drawing_object)
{
    char name[LXW_OBJ_NAME_LENGTH];
    lxw_snprintf(name, LXW_OBJ_NAME_LENGTH, "%s %d", object_name, index);

    xml_attribute_list attributes = {
        {"id", std::to_string(index + 1)},
        {"name", name}
    };

    if (drawing_object)
        attributes.push_back(std::make_pair("descr", drawing_object->description));

    lxw_xml_empty_tag("xdr:cNvPr", attributes);
}

/*
 * Write the <a:picLocks> element.
 */
void drawing::_write_a_pic_locks()
{
    xml_attribute_list attributes = {
        { "noChangeAspect", "1" }
    };

    lxw_xml_empty_tag("a:picLocks", attributes);
}

/*
 * Write the <xdr:cNvPicPr> element.
 */
void drawing::_write_c_nv_pic_pr()
{
    lxw_xml_start_tag("xdr:cNvPicPr");

    /* Write the a:picLocks element. */
    _write_a_pic_locks();

    lxw_xml_end_tag("xdr:cNvPicPr");
}

/*
 * Write the <xdr:nvPicPr> element.
 */
void drawing::_write_nv_pic_pr(uint16_t index, const drawing_object_ptr& drawing_object)
{
    lxw_xml_start_tag("xdr:nvPicPr");

    /* Write the xdr:cNvPr element. */
    _write_c_nv_pr("Picture", index, drawing_object);

    /* Write the xdr:cNvPicPr element. */
    _write_c_nv_pic_pr();

    lxw_xml_end_tag("xdr:nvPicPr");
}

/*
 * Write the <a:blip> element.
 */
void drawing::_write_a_blip(uint16_t index)
{
    char xmlns_r[] = LXW_SCHEMA_OFFICEDOC "/relationships";
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", index);

    xml_attribute_list attributes = {
        {"xmlns:r", xmlns_r},
        {"r:embed", r_id}
    };

    lxw_xml_empty_tag("a:blip", attributes);
}

/*
 * Write the <a:fillRect> element.
 */
void drawing::_write_a_fill_rect()
{
    lxw_xml_empty_tag("a:fillRect");
}

/*
 * Write the <a:stretch> element.
 */
void drawing::_write_a_stretch()
{
    lxw_xml_start_tag("a:stretch");

    /* Write the a:fillRect element. */
    _write_a_fill_rect();

    lxw_xml_end_tag("a:stretch");
}

/*
 * Write the <xdr:blipFill> element.
 */
void drawing::_write_blip_fill(uint16_t index)
{
    lxw_xml_start_tag("xdr:blipFill");

    /* Write the a:blip element. */
    _write_a_blip(index);

    /* Write the a:stretch element. */
    _write_a_stretch();

    lxw_xml_end_tag("xdr:blipFill");
}

/*
 * Write the <a:ext> element.
 */
void drawing::_write_a_ext(const drawing_object_ptr& drawing_object)
{
    xml_attribute_list attributes = {
        {"cx", std::to_string(drawing_object->width)},
        {"cy", std::to_string(drawing_object->height)}
    };

    lxw_xml_empty_tag("a:ext", attributes);
}

/*
 * Write the <a:off> element.
 */
void drawing::_write_a_off(const drawing_object_ptr& drawing_object)
{
    xml_attribute_list attributes = {
        { "x", std::to_string(drawing_object->col_absolute)},
        {"y", std::to_string(drawing_object->row_absolute)}
    };

    lxw_xml_empty_tag("a:off", attributes);
}

/*
 * Write the <a:xfrm> element.
 */
void drawing::_write_a_xfrm(const drawing_object_ptr& drawing_object)
{
    lxw_xml_start_tag("a:xfrm");

    /* Write the a:off element. */
    _write_a_off(drawing_object);

    /* Write the a:ext element. */
    _write_a_ext(drawing_object);

    lxw_xml_end_tag("a:xfrm");
}

/*
 * Write the <a:avLst> element.
 */
void drawing::_write_a_av_lst()
{
    lxw_xml_empty_tag("a:avLst");
}

/*
 * Write the <a:prstGeom> element.
 */
void drawing::_write_a_prst_geom()
{
    xml_attribute_list attributes = {
        {"prst", "rect"}
    };

    lxw_xml_start_tag("a:prstGeom", attributes);

    /* Write the a:avLst element. */
    _write_a_av_lst();

    lxw_xml_end_tag("a:prstGeom");
}

/*
 * Write the <xdr:spPr> element.
 */
void drawing::_write_sp_pr(const drawing_object_ptr& drawing_object)
{
    lxw_xml_start_tag("xdr:spPr");

    /* Write the a:xfrm element. */
    _write_a_xfrm(drawing_object);

    /* Write the a:prstGeom element. */
    _write_a_prst_geom();

    lxw_xml_end_tag("xdr:spPr");
}

/*
 * Write the <xdr:pic> element.
 */
void drawing::_write_pic(uint16_t index, const drawing_object_ptr& drawing_object)
{
    lxw_xml_start_tag("xdr:pic");

    /* Write the xdr:nvPicPr element. */
    _write_nv_pic_pr(index, drawing_object);

    /* Write the xdr:blipFill element. */
    _write_blip_fill(index);

    /* Write the xdr:spPr element. */
    _write_sp_pr(drawing_object);

    lxw_xml_end_tag("xdr:pic");
}

/*
 * Write the <xdr:clientData> element.
 */
void drawing::_write_client_data()
{
    lxw_xml_empty_tag("xdr:clientData");
}

/*
 * Write the <xdr:cNvGraphicFramePr> element.
 */
void drawing::_write_c_nv_graphic_frame_pr()
{
    lxw_xml_empty_tag("xdr:cNvGraphicFramePr");
}

/*
 * Write the <xdr:nvGraphicFramePr> element.
 */
void drawing::_write_nv_graphic_frame_pr(uint16_t index)
{
    lxw_xml_start_tag("xdr:nvGraphicFramePr");

    /* Write the xdr:cNvPr element. */
    _write_c_nv_pr("Chart", index, NULL);

    /* Write the xdr:cNvGraphicFramePr element. */
    _write_c_nv_graphic_frame_pr();

    lxw_xml_end_tag("xdr:nvGraphicFramePr");
}

/*
 * Write the <a:off> element.
 */
void drawing::_write_xfrm_offset()
{
    xml_attribute_list attributes = {
        {"x", "0"},
        {"y", "0"}
    };

    lxw_xml_empty_tag("a:off", attributes);
}

/*
 * Write the <a:ext> element.
 */
void drawing::_write_xfrm_extension()
{
    xml_attribute_list attributes = {
        {"cx", "0"},
        {"cy", "0"}
    };

    lxw_xml_empty_tag("a:ext", attributes);
}

/*
 * Write the <xdr:xfrm> element.
 */
void drawing::_write_xfrm()
{
    lxw_xml_start_tag("xdr:xfrm");

    /* Write the a:off element. */
    _write_xfrm_offset();

    /* Write the a:ext element. */
    _write_xfrm_extension();

    lxw_xml_end_tag("xdr:xfrm");
}

/*
 * Write the <c:chart> element.
 */
void drawing::_write_chart(uint16_t index)
{
    char xmlns_c[] = LXW_SCHEMA_DRAWING "/chart";
    char xmlns_r[] = LXW_SCHEMA_OFFICEDOC "/relationships";
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH];

    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", index);

    xml_attribute_list attributes = {
        {"xmlns:c", xmlns_c},
        {"xmlns:r", xmlns_r},
        {"r:id", r_id}
    };

    lxw_xml_empty_tag("c:chart", attributes);
}

/*
 * Write the <a:graphicData> element.
 */
void drawing::_write_a_graphic_data(uint16_t index)
{
    char uri[] = LXW_SCHEMA_DRAWING "/chart";
    xml_attribute_list attributes = {
        {"uri", uri}
    };

    lxw_xml_start_tag("a:graphicData", attributes);

    /* Write the c:chart element. */
    _write_chart(index);

    lxw_xml_end_tag("a:graphicData");
}

/*
 * Write the <a:graphic> element.
 */
void drawing::_write_a_graphic(uint16_t index)
{

    lxw_xml_start_tag("a:graphic");

    /* Write the a:graphicData element. */
    drawing::_write_a_graphic_data(index);

    lxw_xml_end_tag("a:graphic");
}

/*
 * Write the <xdr:graphicFrame> element.
 */
void drawing::_write_graphic_frame(uint16_t index)
{
    xml_attribute_list attributes {
        {"macro", ""}
    };

    lxw_xml_start_tag("xdr:graphicFrame", attributes);

    /* Write the xdr:nvGraphicFramePr element. */
    _write_nv_graphic_frame_pr(index);

    /* Write the xdr:xfrm element. */
    _write_xfrm();

    /* Write the a:graphic element. */
    _write_a_graphic(index);

    lxw_xml_end_tag("xdr:graphicFrame");
}

/*
 * Write the <xdr:twoCellAnchor> element.
 */
void drawing::_write_two_cell_anchor(uint16_t index, const drawing_object_ptr& drawing_object)
{
    xml_attribute_list attributes;

    if (drawing_object->anchor_type == LXW_ANCHOR_TYPE_IMAGE) {

        if (drawing_object->edit_as == LXW_ANCHOR_EDIT_AS_ABSOLUTE)
            attributes.push_back(std::make_pair("editAs", "absolute"));
        else if (drawing_object->edit_as != LXW_ANCHOR_EDIT_AS_RELATIVE)
            attributes.push_back(std::make_pair("editAs", "oneCell"));
    }
	else if (drawing_object->anchor_type == LXW_ANCHOR_TYPE_CHART)
	{
		if (drawing_object->edit_as == LXW_ANCHOR_EDIT_AS_ABSOLUTE)
            attributes.push_back(std::make_pair("editAs", "absolute"));
	}

    lxw_xml_start_tag("xdr:twoCellAnchor", attributes);

    _write_from(&drawing_object->from);
    _write_to(&drawing_object->to);

    if (drawing_object->anchor_type == LXW_ANCHOR_TYPE_CHART) {
        /* Write the xdr:graphicFrame element for charts. */
        _write_graphic_frame(index);
    }
    else if (drawing_object->anchor_type == LXW_ANCHOR_TYPE_IMAGE) {
        /* Write the xdr:pic element. */
        _write_pic(index, drawing_object);
    }
    else {
        /* Write the xdr:sp element for shapes. */
        /* _drawing_write_sp(self, index, col_absolute, row_absolute, width,
           height,  shape); */
    }

    /* Write the xdr:clientData element. */
    _write_client_data();

    lxw_xml_end_tag("xdr:twoCellAnchor");
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Assemble and write the XML file.
 */
void drawing::assemble_xml_file()
{
    uint16_t index;

    /* Write the XML declaration. */
    _xml_declaration();

    /* Write the xdr:wsDr element. */
    _write_drawing_workspace();

    if (embedded) {
        index = 1;

        for (const auto& drawing_object : drawing_objects) {
            _write_two_cell_anchor(index++, drawing_object);
        }
    }

    lxw_xml_end_tag("xdr:wsDr");
}

}
