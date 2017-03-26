/*****************************************************************************
 * shared_strings - A library for creating Excel XLSX sst files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <xlsxwriter/xmlwriter.hpp>
#include <xlsxwriter/shared_strings.hpp>
#include <xlsxwriter/utility.hpp>
#include <ctype.h>
#include <cstring>


namespace xlsxwriter {

/*
 * Forward declarations.
 */

//STATIC int _element_cmp(struct sst_element *element1,
//                        struct sst_element *element2);

//LXW_RB_GENERATE_ELEMENT(sst_rb_tree, sst_element, sst_tree_pointers,
//                        _element_cmp);

///*****************************************************************************
// *
// * Private functions.
// *
// ****************************************************************************/

///*
// * Comparator for the element structure
// */
//STATIC int
//_element_cmp(struct sst_element *element1, struct sst_element *element2)
//{
//    return strcmp(element1->string, element2->string);
//}

/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/
/*
 * Write the XML declaration.
 */
void sst::_xml_declaration()
{
    lxw_xml_declaration();
}

/*
 * Write the <t> element.
 */
void sst::_write_t(const std::string& string)
{
    xml_attribute_list attributes;

    /* Add attribute to preserve leading or trailing whitespace. */
    if (std::isspace(string[0])
        || std::isspace(string.back()))
        attributes.push_back(std::make_pair("xml:space", "preserve"));

    lxw_xml_data_element("t", string, attributes);
}

/*
 * Write the <si> element.
 */
void sst::_write_si(const std::string& string)
{
    uint8_t escaped_string = false;

    lxw_xml_start_tag("si");

    /* Look for and escape control chars in the string. */
    if (std::strpbrk(string.c_str(), "\x01\x02\x03\x04\x05\x06\x07\x08\x0B\x0C"
                "\x0D\x0E\x0F\x10\x11\x12\x13\x14\x15\x16"
                "\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F")) {
        escaped_string = true;
        _write_t(lxw_escape_control_characters(string));
    }
    else {
        /* Write the t element. */
        _write_t(string);
    }

    lxw_xml_end_tag("si");
}

/*
 * Write the <sst> element.
 */
void sst::_write_sst()
{
    char xmlns[] =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    xml_attribute_list attributes = {
        {"xmlns", xmlns },
        {"count", std::to_string(string_count)},
        {"uniqueCount", std::to_string(unique_count)}
    };
    lxw_xml_start_tag("sst", attributes);
}

/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/

/*
 * Assemble and write the XML file.
 */
void sst::_write_sst_strings()
{
    for (const auto& sst_element : order_list) {
        /* Write the si element. */
        _write_si(sst_element->string);
    }
}

/*
 * Assemble and write the XML file.
 */
void sst::assemble_xml_file()
{
    /* Write the XML declaration. */
    _xml_declaration();

    /* Write the sst element. */
    _write_sst();

    /* Write the sst strings. */
    _write_sst_strings();

    /* Close the sst tag. */
    lxw_xml_end_tag("sst");
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/
/*
 * Add to or find a string in the SST SharedString table and return it's index.
 */
sst_element *sst::get_sst_index(const std::string& string)
{
    std::shared_ptr<sst_element> element = std::make_shared<sst_element>();

    /* Create potential new element with the string and its index. */
    element->index = unique_count;
    element->string = string;

    /* Try to insert it and see whether we already have that string. */
    //existing_element = RB_INSERT(sst_rb_tree, sst->rb_tree, element);

    /* If existing_element is not NULL, then it already existed. */
    /* Free new created element. */
    auto it = strings.insert(element) ;
    if (it.second == false) {
        string_count++;
        return it.first->get();
    }

    /* If it didn't exist, also add it to the insertion order linked list. */
    order_list.push_back(element);

    /* Update SST string counts. */
    string_count++;
    unique_count++;
    return element.get();
}

} // namespace xlsxwriter
