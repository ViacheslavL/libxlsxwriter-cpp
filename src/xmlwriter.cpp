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
void xlsxwriter::lxw_xml_declaration()
{
    fprintf(file, "<?xml version=\"1.0\" "
            "encoding=\"UTF-8\" standalone=\"yes\"?>\n");
}

/*
 * Write an XML start tag with optional attributes.
 */
void xlsxwriter::lxw_xml_start_tag(const char *tag, const std::list<std::pair<std::string, std::string>>& attributes)
{
    fprintf(file, "<%s", tag);

    _fprint_escaped_attributes(file, attributes);

    fprintf(file, ">");
}

/*
 * Write an XML start tag with optional, unencoded, attributes.
 * This is a minor speed optimization for elements that don't need encoding.
 */
void xlsxwriter::lxw_xml_start_tag_unencoded(const char *tag, const std::list<std::pair<std::string, std::string>>& attributes)
{
    struct xml_attribute *attribute;

    fprintf(file, "<%s", tag);

    if (attributes) {
        STAILQ_FOREACH(attribute, attributes, list_entries) {
            fprintf(file, " %s=\"%s\"", attribute->key, attribute->value);
        }
    }

    fprintf(file, ">");
}

/*
 * Write an XML end tag.
 */
void xlsxwriter::lxw_xml_end_tag(const char *tag)
{
    fprintf(file, "</%s>", tag);
}

/*
 * Write an empty XML tag with optional attributes.
 */
void xlsxwriter::lxw_xml_empty_tag(const char *tag, const std::list<std::pair<std::string, std::string>>& attributes)
{
    fprintf(file, "<%s", tag);

    _fprint_escaped_attributes(file, attributes);

    fprintf(file, "/>");
}

/*
 * Write an XML start tag with optional, unencoded, attributes.
 * This is a minor speed optimization for elements that don't need encoding.
 */
void xlsxwriter::lxw_xml_empty_tag_unencoded(const char *tag, const std::list<std::pair<std::string, std::string>>& attributes)
{
    struct xml_attribute *attribute;

    fprintf(file, "<%s", tag);

    if (attributes) {
        STAILQ_FOREACH(attribute, attributes, list_entries) {
            fprintf(file, " %s=\"%s\"", attribute->key, attribute->value);
        }
    }

    fprintf(file, "/>");
}

/*
 * Write an XML element containing data with optional attributes.
 */
void xlsxwriter::lxw_xml_data_element(const std::string& tag, const std::string& data, const std::list<std::pair<std::string, std::string>>& attributes)
{
    fprintf(file, "<%s", tag);

    _fprint_escaped_attributes(file, attributes);

    fprintf(file, ">");

    _fprint_escaped_data(file, data);

    fprintf(file, "</%s>", tag);
}

/*
 * Escape XML characters in attributes.
 */
char * xlsxwriter::_escape_attributes(struct xml_attribute *attribute)
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
char * xlsxwriter::lxw_escape_data(const char *data)
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
char * xlsxwriter::lxw_escape_control_characters(const char *string)
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
void xlsxwriter::_fprint_escaped_attributes(const std::list<std::pair<std::string, std::string>>& attributes)
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
void xlsxwriter::_fprint_escaped_data(const char *data)
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

/* Create a new string XML attribute. */
xml_attribute * xlsxwriter::lxw_new_attribute_str(const char *key, const char *value)
{
    struct xml_attribute *attribute = malloc(sizeof(struct xml_attribute));

    LXW_ATTRIBUTE_COPY(attribute->key, key);
    LXW_ATTRIBUTE_COPY(attribute->value, value);

    return attribute;
}

/* Create a new integer XML attribute. */
xml_attribute * xlsxwriter::lxw_new_attribute_int(const char *key, uint32_t value)
{
    struct xml_attribute *attribute = malloc(sizeof(struct xml_attribute));

    LXW_ATTRIBUTE_COPY(attribute->key, key);
    lxw_snprintf(attribute->value, LXW_MAX_ATTRIBUTE_LENGTH, "%d", value);

    return attribute;
}

/* Create a new double XML attribute. */
xml_attribute * xlsxwriter::lxw_new_attribute_dbl(const char *key, double value)
{
    struct xml_attribute *attribute = malloc(sizeof(struct xml_attribute));

    LXW_ATTRIBUTE_COPY(attribute->key, key);
    lxw_snprintf(attribute->value, LXW_MAX_ATTRIBUTE_LENGTH, "%.16g", value);

    return attribute;
}

} // namespace xlsxwriter
