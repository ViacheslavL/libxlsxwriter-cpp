/*****************************************************************************
 * core - A library for creating Excel XLSX core files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <xlsxwriter/xmlwriter.hpp>
#include <xlsxwriter/core.hpp>
#include <xlsxwriter/utility.hpp>

namespace xlsxwriter {

/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/



/*
 * Convert a time_t struct to a ISO 8601 style "2010-01-01T00:00:00Z" date.
 */
static void
_localtime_to_iso8601_date(time_t *timer, char *str, size_t size)
{
    struct tm *tmp_localtime = nullptr;
    time_t current_time = time(NULL);

    if (*timer)
        tmp_localtime = localtime(timer);
    else
        tmp_localtime = localtime(&current_time);

    strftime(str, size - 1, "%Y-%m-%dT%H:%M:%SZ", tmp_localtime);
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
void core::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <cp:coreProperties> element.
 */
void core::_write_cp_core_properties()
{
    xml_attribute_list attributes = {
        { "xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"},
        {"xmlns:dc", "http://purl.org/dc/elements/1.1/"},
        {"xmlns:dcterms", "http://purl.org/dc/terms/"},
        {"xmlns:dcmitype", "http://purl.org/dc/dcmitype/"},
        {"xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"}
    };

    lxw_xml_start_tag("cp:coreProperties", attributes);
}

/*
 * Write the <dc:creator> element.
 */
void core::_write_dc_creator()
{
    if (!properties->author.empty()) {
        lxw_xml_data_element("dc:creator", properties->author);
    }
    else {
        lxw_xml_data_element("dc:creator", "");
    }
}

/*
 * Write the <cp:lastModifiedBy> element.
 */
void core::_write_cp_last_modified_by()
{
    if (!properties->author.empty()) {
        lxw_xml_data_element("cp:lastModifiedBy", properties->author);
    }
    else {
        lxw_xml_data_element("cp:lastModifiedBy", "");
    }
}

/*
 * Write the <dcterms:created> element.
 */
void core::_write_dcterms_created()
{
    char datetime[LXW_ATTR_32];

    _localtime_to_iso8601_date(&properties->created, datetime,
                               LXW_ATTR_32);

    xml_attribute_list attributes = {
        {"xsi:type", "dcterms:W3CDTF"}
    };

    lxw_xml_data_element("dcterms:created", datetime, attributes);
}

/*
 * Write the <dcterms:modified> element.
 */
void core::_write_dcterms_modified()
{
    char datetime[LXW_ATTR_32];

    _localtime_to_iso8601_date(&properties->created, datetime,
                               LXW_ATTR_32);

    xml_attribute_list attributes = {
        {"xsi:type", "dcterms:W3CDTF"}
    };

    lxw_xml_data_element("dcterms:modified", datetime, attributes);
}

/*
 * Write the <dc:title> element.
 */
void core::_write_dc_title()
{
    if (properties->title.empty())
        return;

    lxw_xml_data_element("dc:title", properties->title);
}

/*
 * Write the <dc:subject> element.
 */
void core::_write_dc_subject()
{
    if (properties->subject.empty())
        return;

    lxw_xml_data_element("dc:subject", properties->subject);
}

/*
 * Write the <cp:keywords> element.
 */
void core::_write_cp_keywords()
{
    if (properties->keywords.empty())
        return;

    lxw_xml_data_element("cp:keywords", properties->keywords);
}

/*
 * Write the <dc:description> element.
 */
void core::_write_dc_description()
{
    if (properties->comments.empty())
        return;

    lxw_xml_data_element("dc:description", properties->comments);
}

/*
 * Write the <cp:category> element.
 */
void core::_write_cp_category()
{
    if (properties->category.empty())
        return;

    lxw_xml_data_element("cp:category", properties->category);
}

/*
 * Write the <cp:contentStatus> element.
 */
void core::_write_cp_content_status()
{
    if (properties->status.empty())
        return;

    lxw_xml_data_element("cp:contentStatus", properties->status);
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Assemble and write the XML file.
 */
void core::assemble_xml_file()
{
    /* Write the XML declaration. */
    _xml_declaration();

    _write_cp_core_properties();
    _write_dc_title();
    _write_dc_subject();
    _write_dc_creator();
    _write_cp_keywords();
    _write_dc_description();
    _write_cp_last_modified_by();
    _write_dcterms_created();
    _write_dcterms_modified();
    _write_cp_category();
    _write_cp_content_status();

    lxw_xml_end_tag("cp:coreProperties");
}

} // namespace xlsxwriter
