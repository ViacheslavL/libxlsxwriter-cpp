/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * packager - A libxlsxwriter library for creating Excel XLSX packager files.
 *
 */
#ifndef __LXW_PACKAGER_HPP__
#define __LXW_PACKAGER_HPP__

#include <stdint.h>
#include "xlsxwriter/third_party/zip.h"

#include <string>

#include "common.hpp"
#include "workbook.hpp"
#include "worksheet.hpp"
#include "shared_strings.hpp"
#include "app.hpp"
#include "core.hpp"
#include "custom.hpp"
#include "theme.hpp"
#include "styles.hpp"
#include "format.hpp"
#include "content_types.hpp"
#include "relationships.hpp"

#define LXW_ZIP_BUFFER_SIZE (16384)

/*  * If zlib returns Z_ERRNO then errno is set and we can trap that. Otherwise
 * return a default libxlsxwriter error. */
#define RETURN_ON_ZIP_ERROR(err, default_err)   \
    if (err == Z_ERRNO)                         \
        return LXW_ERROR_ZIP_FILE_OPERATION;    \
    else                                        \
        return default_err;


namespace xlsxwriter {

class no_zip_file_exception : public std::exception {
};

/*
 * Struct to represent a packager.
 */
class packager {
    friend class xlsxwriter::workbook;
public:
    packager(const std::string& filename, const std::string& tmpdir = std::string());

    uint8_t create_package();

private:

    FILE *file;
    workbook_ptr workbook;

    size_t buffer_size;
    zipFile zipfile;
    zip_fileinfo zipfile_info;
    char *filename;
    char *buffer;
    char *tmpdir;

    uint16_t chart_count;
    uint16_t drawing_count;

    uint8_t _write_workbook_file();
    uint8_t _write_worksheet_files();
    uint8_t _write_image_files();
};

} // namespace xlsxwriter

#endif /* __LXW_PACKAGER_H__ */
