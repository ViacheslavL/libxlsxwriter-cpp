/*****************************************************************************
 * content_types - A library for creating Excel XLSX content_types files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <xlsxwriter/xmlwriter.hpp>
#include <xlsxwriter/content_types.hpp>
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
 * Create a new content_types object.
 */
content_types::content_types()
{
    add_default("rels", LXW_APP_PACKAGE "relationships+xml");
    add_default("xml", "application/xml");

    add_override("/docProps/app.xml", LXW_APP_DOCUMENT "extended-properties+xml");
    add_override("/docProps/core.xml", LXW_APP_PACKAGE "core-properties+xml");
    add_override("/xl/styles.xml", LXW_APP_DOCUMENT "spreadsheetml.styles+xml");
    add_override("/xl/theme/theme1.xml", LXW_APP_DOCUMENT "theme+xml");
    add_override("/xl/workbook.xml", LXW_APP_DOCUMENT "spreadsheetml.sheet.main+xml");
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */
void content_types::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <Types> element.
 */
void content_types::_write_types()
{
    xml_attribute_list attributes = {
        {"xmlns", LXW_SCHEMA_CONTENT}
    };

    lxw_xml_start_tag("Types", attributes);
}

/*
 * Write the <Default> element.
 */
void content_types::_write_default(const std::string& ext, const std::string& type)
{
    xml_attribute_list attributes = {
        {"Extension", ext},
        {"ContentType", type}
    };

    lxw_xml_empty_tag("Default", attributes);
}

/*
 * Write the <Override> element.
 */
void content_types::_write_override(const std::string& part_name, const std::string& type)
{
    xml_attribute_list attributes = {
        {"PartName", part_name},
        {"ContentType", type}
    };

    lxw_xml_empty_tag("Override", attributes);
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Write out all of the <Default> types.
 */
void content_types::_write_defaults()
{
    for(const auto& pair : default_types) {
        _write_default(pair.first, pair.second);
    }
}

/*
 * Write out all of the <Override> types.
 */
void content_types::_write_overrides()
{
    for(const auto& pair : overrides) {
        _write_override(pair.first, pair.second);
    }
}

/*
 * Assemble and write the XML file.
 */
void content_types::assemble_xml_file()
{
    /* Write the XML declaration. */
    _xml_declaration();

    _write_types();
    _write_defaults();
    _write_overrides();

    /* Close the content_types tag. */
    lxw_xml_end_tag("Types");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
/*
 * Add elements to the ContentTypes defaults.
 */
void content_types::add_default(const std::string& key, const std::string& value)
{
    if (key.empty())
        return;

    default_types.push_back(std::make_pair(key, value));
}

/*
 * Add elements to the ContentTypes overrides.
 */
void content_types::add_override(const std::string& key, const std::string& value)
{
    if (key.empty())
        return;

    overrides.push_back(std::make_pair(key, value));
}

/*
 * Add the name of a worksheet to the ContentTypes overrides.
 */
void content_types::add_worksheet_name(const std::string& name)
{
    add_override(name, LXW_APP_DOCUMENT "spreadsheetml.worksheet+xml");
}

/*
 * Add the name of a chart to the ContentTypes overrides.
 */
void content_types::add_chart_name(const std::string& name)
{
    add_override(name, LXW_APP_DOCUMENT "drawingml.chart+xml");
}

/*
 * Add the name of a drawing to the ContentTypes overrides.
 */
void content_types::add_drawing_name(const std::string& name)
{
    add_override(name, LXW_APP_DOCUMENT "drawing+xml");
}

/*
 * Add the sharedStrings link to the ContentTypes overrides.
 */
void content_types::add_shared_strings()
{
    add_override("/xl/sharedStrings.xml", LXW_APP_DOCUMENT "spreadsheetml.sharedStrings+xml");
}

/*
 * Add the calcChain link to the ContentTypes overrides.
 */
void content_types::add_calc_chain()
{
    add_override("/xl/calcChain.xml", LXW_APP_DOCUMENT "spreadsheetml.calcChain+xml");
}

/*
 * Add the custom properties to the ContentTypes overrides.
 */
void content_types::add_custom_properties()
{
    add_override("/docProps/custom.xml", LXW_APP_DOCUMENT "custom-properties+xml");
}

} // namespace xlsxwriter
