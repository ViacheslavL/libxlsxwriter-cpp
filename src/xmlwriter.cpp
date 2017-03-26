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
#include <xlsxwriter/xmlwriter.hpp>
#include <list>
#include <iomanip>

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
    fprintf(file, "<%s", tag.c_str());

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
        fprintf(file, " %s=\"%s\"", attribute.first.c_str(), attribute.second.c_str());
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
    fprintf(file, "<%s", tag.c_str());

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
        fprintf(file, " %s=\"%s\"", attribute.first.c_str(), attribute.second.c_str());
    }
    fprintf(file, "/>");
}

/*
 * Write an XML element containing data with optional attributes.
 */
void xmlwriter::lxw_xml_data_element(const std::string& tag, const std::string& data, const std::list<std::pair<std::string, std::string>>& attributes)
{
    fprintf(file, "<%s", tag.c_str());

    _fprint_escaped_attributes(attributes);

    fprintf(file, ">");

    _fprint_escaped_data(data);

    fprintf(file, "</%s>", tag.c_str());
}

/*
 * Escape XML characters in attributes.
 */
std::string xmlwriter::_escape_attributes(const std::pair<std::string, std::string>& attribute)
{
    std::string encoded;
    encoded.reserve(LXW_MAX_ENCODED_ATTRIBUTE_LENGTH);

    for (const auto& ch : attribute.second) {
        switch (ch) {
            case '&':
                encoded += LXW_AMP;
                break;
            case '<':
                encoded += LXW_LT;
                break;
            case '>':
                encoded += LXW_GT;
                break;
            case '"':
                encoded += LXW_QUOT;
                break;
            default:
                encoded += ch;
                break;
        }
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
    size_t encoded_len = (data.size() * 5 + 1);
    std::string encoded;
    encoded.reserve(encoded_len);

    for (const auto& ch : data) {
        switch (ch) {
            case '&':
                encoded += LXW_AMP;
                break;
            case '<':
                encoded += LXW_LT;
                break;
            case '>':
                encoded += LXW_GT;
                break;
            default:
                encoded += ch;
                break;
        }
    }

    return encoded;
}

/*
 * Escape control characters in strings with with _xHHHH_.
 */
std::string xmlwriter::lxw_escape_control_characters(const std::string& string)
{
    size_t escape_len = sizeof("_xHHHH_") - 1;
    size_t encoded_len = ((string.size()) * escape_len + 1);

    std::string encoded;
    encoded.reserve(encoded_len);

    /*
    char *encoded = (char *) calloc(encoded_len, 1);
    char *p_encoded = encoded;
    */

    for(char ch : string) {
        switch (ch) {
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
            {
                using namespace std;
                stringstream ss;
                ss << "_x" << hex << uppercase << setw(4) << setfill('0') << ch << "_";
                encoded += ss.str();
                break;
            }
            default:
                encoded += ch;
                break;
        }
    }

    return encoded;
}

/* Write out escaped attributes. */
void xmlwriter::_fprint_escaped_attributes(const std::list<std::pair<std::string, std::string>>& attributes)
{
    for (const auto& attribute : attributes) {
        fprintf(file, " %s=", attribute.first.c_str());

        if (!strpbrk(attribute.second.c_str(), "&<>\"")) {
            fprintf(file, "\"%s\"", attribute.second.c_str());
        }
        else {
            std::string encoded = _escape_attributes(attribute);

            if (!encoded.empty()) {
                fprintf(file, "\"%s\"", encoded.c_str());
            }
        }
    }
}

/* Write out escaped XML data. */
void xmlwriter::_fprint_escaped_data(const std::string& data)
{
    /* Escape the data section of the XML element. */
    if (!strpbrk(data.c_str(), "&<>")) {
        fprintf(file, "%s", data.c_str());
    }
    else {
        std::string encoded = lxw_escape_data(data);
        if (!encoded.empty()) {
            fprintf(file, "%s", encoded.c_str());
        }
    }
}

xmlwriter::~xmlwriter()
{

}

} // namespace xlsxwriter
