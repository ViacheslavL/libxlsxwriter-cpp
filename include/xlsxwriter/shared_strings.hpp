/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * shared_strings - A libxlsxwriter library for creating Excel XLSX
 *                  sst files.
 *
 */
#ifndef __LXW_SST_H__
#define __LXW_SST_H__

#include <string.h>
#include <stdint.h>

#include "common.hpp"
#include "xmlwriter.hpp"

#include <unordered_set>
#include <vector>

namespace xlsxwriter {

/* Define a tree.h RB structure for storing shared strings. */
RB_HEAD(sst_rb_tree, sst_element);

/* Define a queue.h structure for storing shared strings in insertion order. */
STAILQ_HEAD(sst_order_list, sst_element);

/* Wrapper around RB_GENERATE_STATIC from tree.h to avoid unused function
 * warnings and to avoid portability issues with the _unused attribute. */
#define LXW_RB_GENERATE_ELEMENT(name, type, field, cmp) \
    RB_GENERATE_INSERT_COLOR(name, type, field, static) \
    RB_GENERATE_INSERT(name, type, field, cmp, static)  \
    /* Add unused struct to allow adding a semicolon */ \
    struct lxw_rb_generate_element{int unused;}

/*
 * Elements of the SST table. They contain pointers to allow them to
 * be stored in a RB tree and also pointers to track the insertion order
 * in a separate list.
 */
struct sst_element {
    uint32_t index;
    std::string string;

    static bool equals( const std::shared_ptr<sst_element>& lhs, const std::shared_ptr<sst_element>& rhs );
};

typedef std::shared_ptr<sst_element> sst_element_ptr;


struct sst_equal
{
    bool operator()(const sst_element_ptr& a, const sst_element_ptr& b) const { return sst_element::equals(a, b); }
};

class packager;
inline size_t hash(const sst_element_ptr& element)
{
    return std::hash<std::string>()(element->string);
}

/*
 * Struct to represent a sst.
 */
class sst : public xmlwriter {
    friend class packager;
public:
    sst_element *get_sst_index(const std::string& string);
    void assemble_xml_file();

    /* Declarations required for unit testing. */

    void _xml_declaration();

    /// TODO make this private in future
    uint32_t string_count;
private:

    uint32_t unique_count;
    std::unordered_set<sst_element_ptr, std::hash<sst_element_ptr>, sst_equal> strings;
    std::vector<sst_element_ptr> order_list;
    /*


    struct sst_rb_tree *rb_tree;
    */

    void _write_t(const std::string &string);
    void _write_si(const std::string &string);
    void _write_sst();
    void _write_sst_strings();
};

typedef std::shared_ptr<sst> sst_ptr;

} // namespace xlsxwriter

#endif /* __LXW_SST_H__ */
