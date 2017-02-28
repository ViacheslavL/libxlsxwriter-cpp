/*****************************************************************************
 * relationships - A library for creating Excel XLSX relationships files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <string.h>
#include "xmlwriter.hpp"
#include "relationships.hpp"
#include "utility.hpp"

/*
 * Forward declarations.
 */


/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/

namespace xlsxwriter {

/*
 * Write the XML declaration.
 */
void relationships::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <Relationship> element.
 */
void relationships::_write_relationship(const std::string& type, const std::string& target, const std::string& target_mode)
{
    char r_id[LXW_MAX_ATTRIBUTE_LENGTH] = { 0 };

    rel_id++;
    lxw_snprintf(r_id, LXW_ATTR_32, "rId%d", rel_id);

    xml_attribute_list attributes = {
        {"Id", r_id},
        {"Type", type},
        {"Target", target}
    };

    if (target_mode)
        attributes.push_back(std::make_pair("TargetMode", target_mode));

    lxw_xml_empty_tag("Relationship", attributes);
}

/*
 * Write the <Relationships> element.
 */
void relationships::_write_relationships()
{
    xml_attribute_list attributes = {
        {"xmlns", LXW_SCHEMA_PACKAGE}
    };

    lxw_xml_start_tag("Relationships", attributes);

    for (const auto& rel : relations) {
        _write_relationship(rel->type, rel->target, rel->target_mode);
    }
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Assemble and write the XML file.
 */
void relationships::assemble_xml_file()
{
    /* Write the XML declaration. */
    _xml_declaration();

    _write_relationships();

    /* Close the relationships tag. */
    lxw_xml_end_tag("Relationships");
}

/*
 * Add a generic container relationship to XLSX .rels xml files.
 */
void relationships::_add_relationship(
    const std::string& schema,
    const std::string& type,
    const std::string& target,
    const std::string& target_mode)
{
    rel_tuple_ptr relationship = std::make_shared<rel_tuple>();

    if (schema.empty() || type.empty() || target.empty())
        return;

    /* Add the schema to the relationship type. */
    /// @todo reduce to LXW_MAX_ATTRIBUTE_TYPE
    relationship->type = schema + type;

    relationship->target = target;

    if (!target_mode.empty()) {
        relationship->target_mode = target_mode;
    }

    relations.push_back(relationship);
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Add a document relationship to XLSX .rels xml files.
 */
void relationships::add_document(
    const std::string& type,
    const std::string& target)
{
    _add_relationship(LXW_SCHEMA_DOCUMENT, type, target, std::string());
}

/*
 * Add a package relationship to XLSX .rels xml files.
 */
void relationships::add_package(
    const std::string& type,
    const std::string& target)
{
    _add_relationship(LXW_SCHEMA_PACKAGE, type, target, std::string());
}

/*
 * Add a MS schema package relationship to XLSX .rels xml files.
 */
void relationships::add_ms_package(
    const std::string& type,
    const std::string& target)
{
    _add_relationship(LXW_SCHEMA_MS, type, target, std::string());
}

/*
 * Add a worksheet relationship to sheet .rels xml files.
 */
void relationships::add_worksheet(
    const std::string& type,
    const std::string& target,
    const std::string& target_mode)
{
    _add_relationship(LXW_SCHEMA_DOCUMENT, type, target, target_mode);
}

} // namespace xlsxwriter
