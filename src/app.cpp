/*****************************************************************************
 * app - A library for creating Excel XLSX app files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xmlwriter.hpp"
#include "app.hpp"
#include "utility.hpp"

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
 * Create a new app object.
 */
app::app()
{
}

/*
 * Free a app object.
 */
app::~app()
{
}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

/*
 * Write the XML declaration.
 */

void app::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <Properties> element.
 */
void app::_write_properties()
{

    char xmlns[] = LXW_SCHEMA_OFFICEDOC "/extended-properties";
    char xmlns_vt[] = LXW_SCHEMA_OFFICEDOC "/docPropsVTypes";

    xml_attribute_list attributes = {
        {"xmlns" , xmlns},
        {"xmlns:vt", xmlns_vt}
    };

    lxw_xml_start_tag("Properties", attributes);
}

/*
 * Write the <Application> element.
 */
void app::_write_application()
{
    lxw_xml_data_element("Application", "Microsoft Excel");
}

/*
 * Write the <DocSecurity> element.
 */
void app::_write_doc_security()
{
    lxw_xml_data_element("DocSecurity", "0");
}

/*
 * Write the <ScaleCrop> element.
 */
void app::_write_scale_crop()
{
    lxw_xml_data_element("ScaleCrop", "false");
}

/*
 * Write the <vt:lpstr> element.
 */
void app::_write_vt_lpstr(const std::string& str)
{
    lxw_xml_data_element("vt:lpstr", str);
}

/*
 * Write the <vt:i4> element.
 */
void app::_write_vt_i4(const std::string& value)
{
    lxw_xml_data_element("vt:i4", value);
}

/*
 * Write the <vt:variant> element.
 */
void app::_write_vt_variant(const std::string& key, const std::string& value)
{
    /* Write the vt:lpstr element. */
    lxw_xml_start_tag("vt:variant");
    _write_vt_lpstr(key);
    lxw_xml_end_tag("vt:variant");

    /* Write the vt:i4 element. */
    lxw_xml_start_tag("vt:variant");
    _write_vt_i4(value);
    lxw_xml_end_tag("vt:variant");
}

/*
 * Write the <vt:vector> element for the heading pairs.
 */
void app::_write_vt_vector_heading_pairs()
{
    xml_attribute_list attributes = {
        {"size", std::to_string(num_heading_pairs * 2)},
        {"baseType", "variant"}
    };

    lxw_xml_start_tag("vt:vector", attributes);

    for(const auto& heading_pair : heading_pairs) {
        _write_vt_variant(heading_pair.first, heading_pair.second);
    }

    lxw_xml_end_tag("vt:vector");
}

/*
 * Write the <vt:vector> element for the named parts.
 */
void app::_write_vt_vector_lpstr_named_parts()
{
    xml_attribute_list attributes = {
        {"size", std::to_string(num_part_names)},
        {"baseType", "lpstr"}
    };

    lxw_xml_start_tag("vt:vector", attributes);

    for (const auto& part_name : part_names) {
        _write_vt_lpstr(part_name);
    }

    lxw_xml_end_tag("vt:vector");
}

/*
 * Write the <HeadingPairs> element.
 */
void app::_write_heading_pairs()
{
    lxw_xml_start_tag("HeadingPairs");

    /* Write the vt:vector element. */
    _write_vt_vector_heading_pairs();

    lxw_xml_end_tag("HeadingPairs");
}

/*
 * Write the <TitlesOfParts> element.
 */
void app::_write_titles_of_parts()
{
    lxw_xml_start_tag("TitlesOfParts");

    /* Write the vt:vector element. */
    _write_vt_vector_lpstr_named_parts();

    lxw_xml_end_tag("TitlesOfParts");
}

/*
 * Write the <Manager> element.
 */
void app::_write_manager()
{
    if (!properties.manager.empty())
        lxw_xml_data_element("Manager", properties.manager);
}

/*
 * Write the <Company> element.
 */
void app::_write_company()
{

    if (!properties.company.empty())
        lxw_xml_data_element("Company", properties.company);
    else
        lxw_xml_data_element("Company", "");
}

/*
 * Write the <LinksUpToDate> element.
 */
void app::_write_links_up_to_date()
{
    lxw_xml_data_element("LinksUpToDate", "false");
}

/*
 * Write the <SharedDoc> element.
 */
void app::_write_shared_doc()
{
    lxw_xml_data_element("SharedDoc", "false");
}

/*
 * Write the <HyperlinkBase> element.
 */
void app::_write_hyperlink_base()
{
    if (!properties.hyperlink_base.empty())
        lxw_xml_data_element("HyperlinkBase", properties.hyperlink_base);
}

/*
 * Write the <HyperlinksChanged> element.
 */
void app::_write_hyperlinks_changed()
{
    lxw_xml_data_element("HyperlinksChanged", "false");
}

/*
 * Write the <AppVersion> element.
 */
void app::_write_app_version()
{
    lxw_xml_data_element("AppVersion", "12.0000");
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Assemble and write the XML file.
 */
void app::_assemble_xml_file()
{

    /* Write the XML declaration. */
    _xml_declaration();

    _write_properties();
    _write_application();
    _write_doc_security();
    _write_scale_crop();
    _write_heading_pairs();
    _write_titles_of_parts();
    _write_manager();
    _write_company();
    _write_links_up_to_date();
    _write_shared_doc();
    _write_hyperlink_base();
    _write_hyperlinks_changed();
    _write_app_version();

    lxw_xml_end_tag("Properties");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Add the name of a workbook Part such as 'Sheet1' or 'Print_Titles'.
 */
void app::_add_part_name(const std::string& name)
{
    part_names.push_back(name);
}

/*
 * Add the name of a workbook Heading Pair such as 'Worksheets', 'Charts' or
 * 'Named Ranges'.
 */
void app::_add_heading_pair(const std::string& key, const std::string& value)
{
    heading_pairs.insert(std::make_pair(key, value));
}

} // namespace xlsxwriter
