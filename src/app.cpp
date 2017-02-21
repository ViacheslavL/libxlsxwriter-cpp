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
    lxw_xml_declaration(file);
}

/*
 * Write the <Properties> element.
 */
void app::_write_properties()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    char xmlns[] = LXW_SCHEMA_OFFICEDOC "/extended-properties";
    char xmlns_vt[] = LXW_SCHEMA_OFFICEDOC "/docPropsVTypes";

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_STR("xmlns", xmlns);
    LXW_PUSH_ATTRIBUTES_STR("xmlns:vt", xmlns_vt);

    lxw_xml_start_tag(file, "Properties", &attributes);

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <Application> element.
 */
void app::_write_application()
{
    lxw_xml_data_element(file, "Application", "Microsoft Excel", NULL);
}

/*
 * Write the <DocSecurity> element.
 */
void app::_write_doc_security()
{
    lxw_xml_data_element(file, "DocSecurity", "0", NULL);
}

/*
 * Write the <ScaleCrop> element.
 */
void app::_write_scale_crop()
{
    lxw_xml_data_element(file, "ScaleCrop", "false", NULL);
}

/*
 * Write the <vt:lpstr> element.
 */
void app::_write_vt_lpstr(const std::string& str)
{
    lxw_xml_data_element(file, "vt:lpstr", str, NULL);
}

/*
 * Write the <vt:i4> element.
 */
void app::_write_vt_i4(const std::string& value)
{
    lxw_xml_data_element(file, "vt:i4", value, NULL);
}

/*
 * Write the <vt:variant> element.
 */
void app::_write_vt_variant(const std::string& key, const std::string& value)
{
    /* Write the vt:lpstr element. */
    lxw_xml_start_tag(file, "vt:variant", NULL);
    _write_vt_lpstr(key);
    lxw_xml_end_tag(file, "vt:variant");

    /* Write the vt:i4 element. */
    lxw_xml_start_tag(file, "vt:variant", NULL);
    _write_vt_i4(value);
    lxw_xml_end_tag(file, "vt:variant");
}

/*
 * Write the <vt:vector> element for the heading pairs.
 */
void app::_write_vt_vector_heading_pairs()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_heading_pair *heading_pair;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("size", num_heading_pairs * 2);
    LXW_PUSH_ATTRIBUTES_STR("baseType", "variant");

    lxw_xml_start_tag(file, "vt:vector", &attributes);

    for(const auto& heading_pair : heading_pairs) {
        _write_vt_variant(heading_pair.first, heading_pair.second);
    }

    lxw_xml_end_tag(file, "vt:vector");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <vt:vector> element for the named parts.
 */
void app::_write_vt_vector_lpstr_named_parts()
{
    struct xml_attribute_list attributes;
    struct xml_attribute *attribute;
    lxw_part_name *part_name;

    LXW_INIT_ATTRIBUTES();
    LXW_PUSH_ATTRIBUTES_INT("size", num_part_names);
    LXW_PUSH_ATTRIBUTES_STR("baseType", "lpstr");

    lxw_xml_start_tag(file, "vt:vector", &attributes);

    for (const auto& part_name : part_names) {
        _write_vt_lpstr(part_name);
    }

    lxw_xml_end_tag(file, "vt:vector");

    LXW_FREE_ATTRIBUTES();
}

/*
 * Write the <HeadingPairs> element.
 */
void app::_write_heading_pairs()
{
    lxw_xml_start_tag(file, "HeadingPairs", NULL);

    /* Write the vt:vector element. */
    _write_vt_vector_heading_pairs();

    lxw_xml_end_tag(file, "HeadingPairs");
}

/*
 * Write the <TitlesOfParts> element.
 */
void app::_write_titles_of_parts()
{
    lxw_xml_start_tag(file, "TitlesOfParts", NULL);

    /* Write the vt:vector element. */
    _write_vt_vector_lpstr_named_parts();

    lxw_xml_end_tag(file, "TitlesOfParts");
}

/*
 * Write the <Manager> element.
 */
void app::_write_manager()
{
    if (!properties.manager.empty())
        lxw_xml_data_element(file, "Manager", properties.manager, NULL);
}

/*
 * Write the <Company> element.
 */
void app::_write_company()
{

    if (!properties.company.empty())
        lxw_xml_data_element(file, "Company", properties.company, NULL);
    else
        lxw_xml_data_element(file, "Company", "", NULL);
}

/*
 * Write the <LinksUpToDate> element.
 */
void app::_write_links_up_to_date()
{
    lxw_xml_data_element(file, "LinksUpToDate", "false", NULL);
}

/*
 * Write the <SharedDoc> element.
 */
void app::_write_shared_doc()
{
    lxw_xml_data_element(file, "SharedDoc", "false", NULL);
}

/*
 * Write the <HyperlinkBase> element.
 */
void app::_write_hyperlink_base()
{
    if (!properties.hyperlink_base.empty())
        lxw_xml_data_element(file, "HyperlinkBase",
                             properties.hyperlink_base, NULL);
}

/*
 * Write the <HyperlinksChanged> element.
 */
void app::_write_hyperlinks_changed()
{
    lxw_xml_data_element(file, "HyperlinksChanged", "false", NULL);
}

/*
 * Write the <AppVersion> element.
 */
void app::_write_app_version()
{
    lxw_xml_data_element(file, "AppVersion", "12.0000", NULL);
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

    lxw_xml_end_tag(file, "Properties");
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
