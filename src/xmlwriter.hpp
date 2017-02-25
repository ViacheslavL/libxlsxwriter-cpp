/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * xmlwriter - A libxlsxwriter library for creating Excel XLSX
 *             XML files.
 *
 * The xmlwriter library is used to create the XML sub-components files
 * in the Excel XLSX file format.
 *
 * This library is used in preference to a more generic XML library to allow
 * for customization and optimization for the XLSX file format.
 *
 * The xmlwriter functions are only used internally and do not need to be
 * called directly by the end user.
 *
 */
#ifndef __XMLWRITER_HPP__
#define __XMLWRITER_HPP__

#include <stdio.h>
#include <stdlib.h>
#include <stdint.h>
#include "common.hpp"
#include <list>

#include <string>

#define LXW_MAX_ATTRIBUTE_LENGTH 256
#define LXW_ATTR_32              32

#define LXW_ATTRIBUTE_COPY(dst, src)                    \
    do{                                                 \
        strncpy(dst, src, LXW_MAX_ATTRIBUTE_LENGTH -1); \
        dst[LXW_MAX_ATTRIBUTE_LENGTH - 1] = '\0';       \
    } while (0)


/* Attribute used in XML elements. */
struct xml_attribute {
    char key[LXW_MAX_ATTRIBUTE_LENGTH];
    char value[LXW_MAX_ATTRIBUTE_LENGTH];

    /* Make the struct a queue.h list element. */
    STAILQ_ENTRY (xml_attribute) list_entries;
};

/* Use queue.h macros to define the xml_attribute_list type. */
//STAILQ_HEAD(xml_attribute_list, xml_attribute);

/* Create a new attribute struct to add to a xml_attribute_list. */
struct xml_attribute *lxw_new_attribute_str(const char *key,
                                            const char *value);
struct xml_attribute *lxw_new_attribute_int(const char *key, uint32_t value);
struct xml_attribute *lxw_new_attribute_dbl(const char *key, double value);

/* Macro to initialize the xml_attribute_list pointers. */
#define LXW_INIT_ATTRIBUTES()                                 \
    STAILQ_INIT(&attributes)

/* Macro to add attribute string elements to xml_attribute_list. */
#define LXW_PUSH_ATTRIBUTES_STR(key, value)                   \
    do {                                                      \
    attribute = lxw_new_attribute_str((key), (value));        \
    STAILQ_INSERT_TAIL(&attributes, attribute, list_entries); \
    } while (0)

/* Macro to add attribute int values to xml_attribute_list. */
#define LXW_PUSH_ATTRIBUTES_INT(key, value)                   \
    do {                                                      \
    attribute = lxw_new_attribute_int((key), (value));        \
    STAILQ_INSERT_TAIL(&attributes, attribute, list_entries); \
    } while (0)

/* Macro to add attribute double values to xml_attribute_list. */
#define LXW_PUSH_ATTRIBUTES_DBL(key, value)                   \
    do {                                                      \
    attribute = lxw_new_attribute_dbl((key), (value));        \
    STAILQ_INSERT_TAIL(&attributes, attribute, list_entries); \
    } while (0)

/* Macro to free xml_attribute_list and attribute. */
#define LXW_FREE_ATTRIBUTES()                                 \
    while (!STAILQ_EMPTY(&attributes)) {                      \
        attribute = STAILQ_FIRST(&attributes);                \
        STAILQ_REMOVE_HEAD(&attributes, list_entries);        \
        free(attribute);                                      \
    }

namespace xlsxwriter {

typedef std::list<std::pair<std::string, std::string>> xml_attribute_list;

class xmlwriter {
public:
    virtual ~xmlwriter();
protected:

    /**
     * Create the XML declaration in an XML file.
     *
     * @param xmlfile A FILE pointer to the output XML file.
     */
    void lxw_xml_declaration();

    /**
     * Write an XML start tag with optional attributes.
     *
     * @param tag        The XML tag to write.
     * @param attributes An optional list of attributes to add to the tag.
     */
    void lxw_xml_start_tag(const char *tag,
                           const std::list<std::pair<std::string, std::string>>& attributes = xml_attribute_list());

    /**
     * Write an XML start tag with optional un-encoded attributes.
     * This is a minor optimization for attributes that don't need encoding.
     *
     * @param tag        The XML tag to write.
     * @param attributes An optional list of attributes to add to the tag.
     */
    void lxw_xml_start_tag_unencoded(const char *tag, const std::list<std::pair<std::string, std::string>>& attributes);

    /**
     * Write an XML end tag.
     *
     * @param xmlfile    A FILE pointer to the output XML file.
     * @param tag        The XML tag to write.
     */
    void lxw_xml_end_tag(const char *tag);

    /**
     * Write an XML empty tag with optional attributes.
     *
     * @param tag        The XML tag to write.
     * @param attributes An optional list of attributes to add to the tag.
     */
    void lxw_xml_empty_tag(const char *tag, const std::list<std::pair<std::string, std::string>>& attributes = xml_attribute_list());

    /**
     * Write an XML empty tag with optional un-encoded attributes.
     * This is a minor optimization for attributes that don't need encoding.
     *
     * @param tag        The XML tag to write.
     * @param attributes An optional list of attributes to add to the tag.
     */
    void lxw_xml_empty_tag_unencoded(const char *tag,
                                     const std::list<std::pair<std::string, std::string>>& attributes);

    /**
     * Write an XML element containing data and optional attributes.
     *
     * @param tag        The XML tag to write.
     * @param data       The data section of the XML element.
     * @param attributes An optional list of attributes to add to the tag.
     */
    void lxw_xml_data_element(const std::string& tag,
                              const std::string& data,
                              const std::list<std::pair<std::string, std::string>>& attributes = xml_attribute_list());

    char *lxw_escape_control_characters(const char *string);

    char *lxw_escape_data(const char *data);

    FILE* file;
};



} // namespace xlsxwriter


#endif /* __XMLWRITER_HPP__ */
