/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * core - A libxlsxwriter library for creating Excel XLSX core files.
 *
 */
#ifndef __LXW_CORE_H__
#define __LXW_CORE_H__

#include <stdint.h>

#include "workbook.hpp"
#include "common.hpp"

namespace xlsxwriter {

/*
 * Struct to represent a core.
 */
class core : public xmlwriter {
public:
    void assemble_xml_file();

    /* Declarations required for unit testing. */

    void _xml_declaration();

private:

    doc_properties *properties;

    void _write_cp_core_properties();
    void _write_dc_creator();
    void _write_dcterms_modified();
    void _write_dcterms_created();
    void _write_cp_last_modified_by();
    void _write_dc_title();
    void _write_cp_keywords();
    void _write_dc_description();
    void _write_cp_category();
    void _write_cp_content_status();
    void _write_dc_subject();
};

} // namespace xlsxwriter

#endif /* __LXW_CORE_H__ */
