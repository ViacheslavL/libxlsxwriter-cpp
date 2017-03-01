/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * relationships - A libxlsxwriter library for creating Excel XLSX
 *                 relationships files.
 *
 */
#ifndef __LXW_RELATIONSHIPS_HPP__
#define __LXW_RELATIONSHIPS_HPP__

#include <stdint.h>

#include "common.hpp"
#include "xmlwriter.hpp"
#include <vector>

namespace xlsxwriter {

class packager;

struct rel_tuple {
    std::string type;
    std::string target;
    std::string target_mode;
};

typedef std::shared_ptr<rel_tuple> rel_tuple_ptr;

/*
 * Struct to represent a relationships.
 */
class relationships : public xmlwriter{
    friend class packager;
public:
    void assemble_xml_file();

    void add_package(const std::string &type, const std::string &target);
    void add_document(const std::string &type, const std::string &target);
    void add_ms_package(const std::string &type, const std::string &target);
    void add_worksheet(const std::string &type, const std::string &target, const std::string &target_mode);

    /* Declarations required for unit testing. */
    #ifdef TESTING




    #endif /* TESTING */
    void _xml_declaration();

private:
    uint32_t rel_id;
    std::vector<rel_tuple_ptr> relations;


    void _write_relationships();
    void _add_relationship(const std::string &schema, const std::string &type, const std::string &target, const std::string &target_mode);
    void _write_relationship(const std::string &type, const std::string &target, const std::string &target_mode);
};


typedef std::shared_ptr<relationships> relationships_ptr;

} // namespace xlsxwriter

#endif /* __LXW_RELATIONSHIPS_HPP__ */
