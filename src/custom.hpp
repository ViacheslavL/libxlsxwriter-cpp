/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * custom - A libxlsxwriter library for creating Excel custom property files.
 *
 */
#ifndef __LXW_CUSTOM_HPP__
#define __LXW_CUSTOM_HPP__

#include <stdint.h>

#include "common.hpp"
#include "xmlwriter.hpp"

namespace xlsxwriter {


class packager;
/*
 * class to represent a custom property file object.
 */
class custom : public xmlwriter {
    friend class packager;
public:

    custom(const std::list<custom_property_ptr>& properties);

    void assemble_xml_file();

    /* Declarations required for unit testing. */

    void _xml_declaration();
private:

    const std::list<custom_property_ptr>& custom_properties;
    uint32_t pid;
    void _chart_write_vt_lpwstr(char *value);
    void _chart_write_vt_r_8(double value);
    void _write_vt_i_4(int32_t value);
    void _write_vt_filetime(lxw_datetime *datetime);
    void _write_vt_bool(uint8_t value);
    void _write_custom_properties();
    void _chart_write_custom_property(const custom_property_ptr &custom_property);
};

typedef std::shared_ptr<custom> custom_ptr;

}

#endif /* __LXW_CUSTOM_HPP__ */
