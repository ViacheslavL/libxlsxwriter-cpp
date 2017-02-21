/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * app - A libxlsxwriter library for creating Excel XLSX app files.
 *
 */
#ifndef __LXW_APP_HPP__
#define __LXW_APP_HPP__

#include "workbook.hpp"
#include "common.hpp"

#include <map>
#include <list>

namespace xlsxwriter {

/* Class to represent an App object. */
class app {
public:

    app();
    ~app();
    void _assemble_xml_file();
    void _add_part_name(const std::string& name);
    void _add_heading_pair(const std::string& key, const std::string& value);

    /* Declarations required for unit testing. */
    void _xml_declaration();

private:

    FILE *file;

    std::map<std::string, std::string> heading_pairs;
    std::list<std::string> part_names;
    doc_properties properties;

    uint32_t num_heading_pairs;
    uint32_t num_part_names;

    void _write_titles_of_parts();
    void _write_vt_vector_heading_pairs();
    void _write_vt_variant(const std::string &key, const std::string &value);
    void _write_vt_i4(const std::string &value);
    void _write_vt_lpstr(const std::string &str);
    void _write_scale_crop();
    void _write_properties();
    void _write_application();
    void _write_doc_security();
    void _write_vt_vector_lpstr_named_parts();
    void _write_heading_pairs();
    void _write_manager();
    void _write_links_up_to_date();
    void _write_shared_doc();
    void _write_hyperlink_base();
    void _write_hyperlinks_changed();
    void _write_app_version();
    void _app_add_part_name(const std::string &name);
    void _app_add_heading_pair(const std::string &key, const std::string &value);
    void _write_company();
};

} // namespace xlsxwriter

#endif /* __LXW_APP_HPP__ */
