/*****************************************************************************
 * custom - A library for creating Excel custom property files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xmlwriter.hpp"
#include "custom.hpp"
#include "utility.hpp"


/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

namespace xlsxwriter {

/*
 * Assemble and write the XML file.
 */
custom::custom(const std::list<custom_property_ptr> &properties) :
    custom_properties(properties)
{

}

/*
 * Write the XML declaration.
 */
void custom::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <vt:lpwstr> element.
 */
void custom::_chart_write_vt_lpwstr(char *value)
{
    lxw_xml_data_element("vt:lpwstr", value);
}

/*
 * Write the <vt:r8> element.
 */
void custom::_chart_write_vt_r_8(double value)
{
    char data[LXW_ATTR_32];

    lxw_snprintf(data, LXW_ATTR_32, "%.16g", value);

    lxw_xml_data_element("vt:r8", data);
}

/*
 * Write the <vt:i4> element.
 */
void custom::_write_vt_i_4(int32_t value)
{
    char data[LXW_ATTR_32];

    lxw_snprintf(data, LXW_ATTR_32, "%d", value);

    lxw_xml_data_element("vt:i4", data);
}

/*
 * Write the <vt:bool> element.
 */
void custom::_write_vt_bool(uint8_t value)
{
    if (value)
        lxw_xml_data_element("vt:bool", "true");
    else
        lxw_xml_data_element("vt:bool", "false");
}

/*
 * Write the <vt:filetime> element.
 */
void custom::_write_vt_filetime(lxw_datetime *datetime)
{
    char data[LXW_DATETIME_LENGTH];

    lxw_snprintf(data, LXW_DATETIME_LENGTH, "%4d-%02d-%02dT%02d:%02d:%02dZ",
                 datetime->year, datetime->month, datetime->day,
                 datetime->hour, datetime->min, (int) datetime->sec);

    lxw_xml_data_element("vt:filetime", data);
}

/*
 * Write the <property> element.
 */
void custom::_chart_write_custom_property(const custom_property_ptr& custom_property)
{
    char fmtid[] = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    pid++;

    xml_attribute_list attributes = {
        {"fmtid", fmtid},
        {"pid", std::to_string(pid + 1)},
        {"name", custom_property->name}
    };

    lxw_xml_start_tag("property", attributes);

    if (custom_property->type == LXW_CUSTOM_STRING) {
        /* Write the vt:lpwstr element. */
        _chart_write_vt_lpwstr(custom_property->u.string);
    }
    else if (custom_property->type == LXW_CUSTOM_DOUBLE) {
        /* Write the vt:r8 element. */
        _chart_write_vt_r_8(custom_property->u.number);
    }
    else if (custom_property->type == LXW_CUSTOM_INTEGER) {
        /* Write the vt:i4 element. */
        _write_vt_i_4(custom_property->u.integer);
    }
    else if (custom_property->type == LXW_CUSTOM_BOOLEAN) {
        /* Write the vt:bool element. */
        _write_vt_bool(custom_property->u.boolean);
    }
    else if (custom_property->type == LXW_CUSTOM_DATETIME) {
        /* Write the vt:filetime element. */
        _write_vt_filetime(&custom_property->u.datetime);
    }

    lxw_xml_end_tag("property");
}

/*
 * Write the <Properties> element.
 */
void custom::_write_custom_properties()
{
    char xmlns[] = LXW_SCHEMA_OFFICEDOC "/custom-properties";
    char xmlns_vt[] = LXW_SCHEMA_OFFICEDOC "/docPropsVTypes";

    xml_attribute_list attributes = {
        { "xmlns", xmlns},
        {"xmlns:vt", xmlns_vt}
    };

    lxw_xml_start_tag("Properties", attributes);

    for (const auto& custom_property : custom_properties) {
        _chart_write_custom_property(custom_property);
    }
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/


void custom::assemble_xml_file()
{
    /* Write the XML declaration. */
    _xml_declaration();

    _write_custom_properties();

    lxw_xml_end_tag("Properties");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

} // namespace xlsxwriter
