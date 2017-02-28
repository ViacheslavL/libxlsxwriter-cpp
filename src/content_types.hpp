/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * content_types - A libxlsxwriter library for creating Excel XLSX
 *                 content_types files.
 *
 */
#ifndef __LXW_CONTENT_TYPES_H__
#define __LXW_CONTENT_TYPES_H__

#include <stdint.h>
#include <string.h>
#include "xmlwriter.hpp"

#include "common.hpp"

namespace xlsxwriter {

#define LXW_APP_PACKAGE  "application/vnd.openxmlformats-package."
#define LXW_APP_DOCUMENT "application/vnd.openxmlformats-officedocument."

/*
 * Struct to represent a content_types.
 */
struct content_types : public xmlwriter {

public:

    content_types();
    ~content_types();

    void assemble_xml_file();

    void add_default(const std::string& key, const std::string& value);

    void add_override(const std::string& key, const std::string& value);

    void add_worksheet_name(const std::string& name);

    void add_chart_name(const std::string& name);

    void add_drawing_name(const std::string& name);

    void add_shared_strings();

    void add_calc_chain();

    void add_custom_properties();

    /* Declarations required for unit testing. */
    void _content_types_xml_declaration();
    void _write_default(const std::string& ext, const std::string& type);
    void _write_override(const std::string& part_name, const std::string& type);

private:
    std::list<std::pair<std::string, std::string>> default_types;
    std::list<std::pair<std::string, std::string>> overrides;

    void _write_defaults();
    void _write_types();
    void _xml_declaration();
    void _write_overrides();
};

typedef std::shared_ptr<content_types> content_types_ptr;

} // namespace xlsxwriter

#endif /* __LXW_CONTENT_TYPES_H__ */
