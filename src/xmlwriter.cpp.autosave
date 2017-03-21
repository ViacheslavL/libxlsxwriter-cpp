/*****************************************************************************
 * xmlwriter - A base library for libxlsxwriter libraries.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include "xmlwriter.hpp"
#include <list>

#define LXW_AMP  "&amp;"
#define LXW_LT   "&lt;"
#define LXW_GT   "&gt;"
#define LXW_QUOT "&quot;"

/* Defines. */
#define LXW_MAX_ENCODED_ATTRIBUTE_LENGTH (LXW_MAX_ATTRIBUTE_LENGTH*6)


namespace xlsxwriter {
/*
 * Write the XML declaration.
 */
void xmlwriter::lxw_xml_declaration()
{
    fprintf(file, "<?xml version=\"1.0\" "
            "encoding=\"UTF-8\" standalone=\"yes\"?>\n");
}

/*
 * Write an XML start tag with optional attributes.
 */
void xmlwriter::lxw_xml_start_tag(const std::string& tag, const std::list<std::pair<std::string, std::string>>& attributes)
{
    fprintf(file, "<%s", tag);

    _fprint_escaped_attributes(attributes);

    fprintf(file, ">");
}

/*
 * Write an XML start tag with optional, unencoded, attributes.
 * This is a minor speed optimization for elements that don't need encoding.
 */
void xmlwriter::lxw_xml_start_tag_unencoded(const char *tag, const std::list<std::pair<std::string, std::string>>& attributes)
{
    fprintf(file, "<%s", tag);

    for (const auto& attribute : attributes) {
        fprintf(file, " %s=\"%s\"", attribute.first, attribute.second);
    }

    fprintf(file, ">");
}

/*
 * Write an XML end tag.
 */
void xmlwriter::lxw_xml_end_tag(const std::string& tag)
{
    fprintf(file, "</%s>", tag.c_str());
}

/*
 * Write an empty XML tag with optional attributes.
 */
void xmlwriter::lxw_xml_empty_tag(const std::string&  tag, const std::list<std::pair<std::string, std::string>>& attributes)
{
    fprintf(file, "<%s", tag);

    _fprint_escaped_attributes(attributes);

    fprintf(file, "/>");
}

/*
 * Write an XML start tag with optional, unencoded, attributes.
 * This is a minor speed optimization for elements that don't need encoding.
 */
void xmlwriter::lxw_xml_empty_tag_unencoded(const char *tag, const std::list<std::pair<std::string, std::string>>& attributes)
{
    fprintf(file, "<%s", tag);
    for (const auto& attribute : attributes) {
        fprintf(file, " %s=\"%s\"", attribute.first, attribute.second);
    }
    fprintf(file, "/>");
}

/*
 * Write an XML element containing data with optional attributes.
 */
void xmlwriter::lxw_xml_data_element(const std::string& tag, const std::string& data, const std::list<std::pair<std::string, std::string>>& attributes)
{
    fprintf(file, "<%s", tag);

    _fprint_escaped_attributes(attributes);

    fprintf(file, ">");

    _fprint_escaped_data(data);

    fprintf(file, "</%s>", tag);
}

/*
 * Escape XML characters in attributes.
 */
char * xmlwriter::_escape_attributes(struct xml_attribute *attribute)
{
    char *encoded = (char *) calloc(LXW_MAX_ENCODED_ATTRIBUTE_LENGTH, 1);
    char *p_encoded = encoded;
    char *p_attr = attribute->value;

    while (*p_attr) {
        switch (*p_attr) {
            case '&':
                strncat(p_encoded, LXW_AMP, sizeof(LXW_AMP) - 1);
                p_encoded += sizeof(LXW_AMP) - 1;
                break;
            case '<':
                strncat(p_encoded, LXW_LT, sizeof(LXW_LT) - 1);
                p_encoded += sizeof(LXW_LT) - 1;
                break;
            case '>':
                strncat(p_encoded, LXW_GT, sizeof(LXW_GT) - 1);
                p_encoded += sizeof(LXW_GT) - 1;
                break;
            case '"':
                strncat(p_encoded, LXW_QUOT, sizeof(LXW_QUOT) - 1);
                p_encoded += sizeof(LXW_QUOT) - 1;
                break;
            default:
                *p_encoded = *p_attr;
                p_encoded++;
                break;
        }
        p_attr++;
    }

    return encoded;
}

/*
 * Escape XML characters in data sections of tags.
 * Note, this is different from _escape_attributes()
 * in that double quotes are not escaped by Excel.
 */
std::string xmlwriter::lxw_escape_data(const std::string& data)
{
    size_t encoded_len = (strlen(data) * 5 + 1);

    char *encoded = (char *) calloc(encoded_len, 1);
    char *p_encoded = encoded;

    while (*data) {
        switch (*data) {
            case '&':
                strncat(p_encoded, LXW_AMP, sizeof(LXW_AMP) - 1);
                p_encoded += sizeof(LXW_AMP) - 1;
                break;
            case '<':
                strncat(p_encoded, LXW_LT, sizeof(LXW_LT) - 1);
                p_encoded += sizeof(LXW_LT) - 1;
                break;
            case '>':
                strncat(p_encoded, LXW_GT, sizeof(LXW_GT) - 1);
                p_encoded += sizeof(LXW_GT) - 1;
                break;
            default:
                *p_encoded = *data;
                p_encoded++;
                break;
        }
        data++;
    }

    return encoded;
}

/*
 * Escape control characters in strings with with _xHHHH_.
 */
std::string xmlwriter::lxw_escape_control_characters(const char *string)
{
    size_t escape_len = sizeof("_xHHHH_") - 1;
    size_t encoded_len = (strlen(string) * escape_len + 1);

    char *encoded = (char *) calloc(encoded_len, 1);
    char *p_encoded = encoded;

    while (*string) {
        switch (*string) {
            case '\x01':
            case '\x02':
            case '\x03':
            case '\x04':
            case '\x05':
            case '\x06':
            case '\x07':
            case '\x08':
            case '\x0B':
            case '\x0C':
            case '\x0D':
            case '\x0E':
            case '\x0F':
            case '\x10':
            case '\x11':
            case '\x12':
            case '\x13':
            case '\x14':
            case '\x15':
            case '\x16':
            case '\x17':
            case '\x18':
            case '\x19':
            case '\x1A':
            case '\x1B':
            case '\x1C':
            case '\x1D':
            case '\x1E':
            case '\x1F':
                lxw_snprintf(p_encoded, escape_len + 1, "_x%04X_", *string);
                p_encoded += escape_len;
                break;
            default:
                *p_encoded = *string;
                p_encoded++;
                break;
        }
        string++;
    }

    return encoded;
}

/* Write out escaped attributes. */
void xmlwriter::_fprint_escaped_attributes(const std::list<std::pair<std::string, std::string>>& attributes)
{
    for (const auto& attribute : attributes) {
        fprintf(file, " %s=", attribute.first);

        if (!strpbrk(attribute->value, "&<>\"")) {
            fprintf(file, "\"%s\"", attribute.second);
        }
        else {
            char *encoded = _escape_attributes(attribute);

            if (encoded) {
                fprintf(file, "\"%s\"", encoded);

                free(encoded);
            }
        }
    }
}

/* Write out escaped XML data. */
void xmlwriter::_fprint_escaped_data(const char *data)
{
    /* Escape the data section of the XML element. */
    if (!strpbrk(data, "&<>")) {
        fprintf(file, "%s", data);
    }
    else {
        char *encoded = lxw_escape_data(data);
        if (encoded) {
            fprintf(file, "%s", encoded);
            free(encoded);
        }
    }
}

xmlwriter::~xmlwriter()
{

}

} // namespace xlsxwriter
