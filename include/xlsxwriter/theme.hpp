/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * theme - A libxlsxwriter library for creating Excel XLSX theme files.
 *
 */
#ifndef __LXW_THEME_H__
#define __LXW_THEME_H__

#include <stdint.h>

#include "common.hpp"
#include "xmlwriter.hpp"

namespace xlsxwriter {

class packager;
/*
 * Struct to represent a theme.
 */
class theme : public xmlwriter {
    friend class packager;
public:
    void xml_declaration();
    void assemble_xml_file();
};

/* Declarations required for unit testing. */
#ifdef TESTING
#endif /* TESTING */


} // namespace xlsxwriter

#endif /* __LXW_THEME_H__ */
