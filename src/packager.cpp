/*****************************************************************************
 * packager - A library for creating Excel XLSX packager files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <xlsxwriter/xmlwriter.hpp>
#include <xlsxwriter/packager.hpp>
#include <xlsxwriter/hash_table.hpp>
#include <xlsxwriter/utility.hpp>

namespace xlsxwriter {

uint8_t _add_file_to_zip(FILE * file, const char *filename);

/*
 * Forward declarations.
 */

/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/
/* Avoid non MSVC definition of _WIN32 in MinGW. */

#ifdef __MINGW32__
#undef _WIN32
#endif

#ifdef _WIN32

/* Silence Windows warning with duplicate symbol for SLIST_ENTRY in local
 * queue.h and widows.h. */
#undef SLIST_ENTRY

#include <windows.h>
#include "../third_party/minizip/iowin32.h"

zipFile
_open_zipfile_win32(const char *filename)
{
    int n;
    zlib_filefunc64_def filefunc;

    wchar_t wide_filename[_MAX_PATH + 1] = L"";

    /* Build a UTF-16 filename for Win32. */
    n = MultiByteToWideChar(CP_UTF8, 0, filename, (int) strlen(filename),
                            wide_filename, _MAX_PATH);

    if (n == 0) {
        LXW_ERROR("MultiByteToWideChar error");
        return NULL;
    }

    /* Use the native Win32 file handling functions with minizip. */
    fill_win32_filefunc64(&filefunc);

    return zipOpen2_64(wide_filename, 0, NULL, &filefunc);
}

#endif

/*
 * Create a new packager object.
 */
packager::packager(const std::string& filename, const std::string& tmpdir)
    : chart_count(0)
    , drawing_count(0)
{
    this->filename = filename;
    this->tmpdir = tmpdir;

    /* Initialize the zip_fileinfo struct to Jan 1 1980 like Excel. */
    zipfile_info.tmz_date.tm_sec = 0;
    zipfile_info.tmz_date.tm_min = 0;
    zipfile_info.tmz_date.tm_hour = 0;
    zipfile_info.tmz_date.tm_mday = 1;
    zipfile_info.tmz_date.tm_mon = 0;
    zipfile_info.tmz_date.tm_year = 1980;
    zipfile_info.dosDate = 0;
    zipfile_info.internal_fa = 0;
    zipfile_info.external_fa = 0;

    /* Create a zip container for the xlsx file. */
#ifdef _WIN32
    zipfile = _open_zipfile_win32(this->filename);
#else
    zipfile = zipOpen(this->filename.c_str(), 0);
#endif

    if (zipfile == NULL)
        throw new no_zip_file_exception();
}

/*****************************************************************************
 *
 * File assembly functions.
 *
 ****************************************************************************/
/*
 * Write the workbook.xml file.
 */
uint8_t packager::_write_workbook_file()
{
    workbook->file = lxw_tmpfile(tmpdir.c_str());
    if (!workbook->file)
        return LXW_ERROR_CREATING_TMPFILE;

    workbook->assemble_xml_file();

    uint8_t err = _add_file_to_zip( workbook->file, "xl/workbook.xml");
    RETURN_ON_ERROR(err);

    fclose(workbook->file);

    return 0;
}

/*
 * Write the worksheet files.
 */
uint8_t packager::_write_worksheet_files()
{
    char sheetname[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;
    int err;

    for (const auto& worksheet : workbook->worksheets) {
        lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
                     "xl/worksheets/sheet%d.xml", index++);

        if (worksheet->optimize_row)
            worksheet->write_single_row();

        worksheet->file = lxw_tmpfile(tmpdir.c_str());
        if (!worksheet->file)
            return LXW_ERROR_CREATING_TMPFILE;

        worksheet->assemble_xml_file();

        err = _add_file_to_zip(worksheet->file, sheetname);
        RETURN_ON_ERROR(err);

        fclose(worksheet->file);
    }

    return 0;
}

/*
 * Write the /xl/media/image?.xml files.
 */
uint8_t packager::_write_image_files()
{
    int err;

    char filename[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;

    for(const auto& sheet : workbook->worksheets) {

        if (sheet->image_data.empty())
            continue;

        for (const auto& image : sheet->image_data) {
            lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                         "xl/media/image%d.%s", index++, image->extension.c_str());

            rewind(image->stream);

            err = _add_file_to_zip(image->stream, filename);
            RETURN_ON_ERROR(err);

            fclose(image->stream);
        }
    }

    return 0;
}

/*
 * Write the chart files.
 */
uint8_t packager::_write_chart_files()
{
    char sheetname[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;
    int err;

    for(const auto& chart: workbook->ordered_charts) {

        lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
                     "xl/charts/chart%d.xml", index++);

        chart->file = lxw_tmpfile(tmpdir.c_str());
        if (!chart->file)
            return LXW_ERROR_CREATING_TMPFILE;

        chart->assemble_xml_file();

        err = _add_file_to_zip(chart->file, sheetname);
        RETURN_ON_ERROR(err);

        chart_count++;

        fclose(chart->file);
    }

    return 0;
}

/*
 * Write the drawing files.
 */
uint8_t packager::_write_drawing_files()
{
    char filename[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;
    int err;

    for(const auto& worksheet : workbook->worksheets) {
        const std::shared_ptr<xlsxwriter::drawing>& drawing = worksheet->drawing;

        if (drawing) {
            lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                         "xl/drawings/drawing%d.xml", index++);

            drawing->file = lxw_tmpfile(tmpdir.c_str());
            if (!drawing->file)
                return LXW_ERROR_CREATING_TMPFILE;

            drawing->assemble_xml_file();
            err = _add_file_to_zip(drawing->file, filename);
            RETURN_ON_ERROR(err);

            fclose(drawing->file);

            drawing_count++;
        }
    }

    return 0;
}

/*
 * Write the sharedStrings.xml file.
 */
uint8_t packager::_write_shared_strings_file()
{
    xlsxwriter::sst *sst = workbook->sst.get();
    int err;

    /* Skip the sharedStrings file if there are no shared strings. */
    if (sst->string_count == 0)
        return 0;

    sst->file = lxw_tmpfile(tmpdir.c_str());
    if (!sst->file)
        return LXW_ERROR_CREATING_TMPFILE;

    sst->assemble_xml_file();

    err = _add_file_to_zip(sst->file, "xl/sharedStrings.xml");
    RETURN_ON_ERROR(err);

    fclose(sst->file);

    return 0;
}

/*
 * Write the app.xml file.
 */
uint8_t packager::_write_app_file()
{
    std::shared_ptr<xlsxwriter::app> app = std::make_shared<xlsxwriter::app>();
    uint16_t named_range_count = 0;
    std::string number;
    int err;

    app->file = lxw_tmpfile(tmpdir.c_str());
    if (!app->file)
        return LXW_ERROR_CREATING_TMPFILE;

    number = std::to_string( workbook->worksheets.size());

    app->add_heading_pair("Worksheets", number);

    for(const auto& worksheet: workbook->worksheets) {
        app->add_part_name(worksheet->name);
    }

    /* Add the Named Ranges parts. */
    for (const auto& defined_name : workbook->defined_names) {
        const char* has_range = strchr(defined_name->formula.c_str(), '!');
        const char* autofilter = strstr(defined_name->app_name.c_str(), "_FilterDatabase");

        /* Only store defined names with ranges (except for autofilters). */
        if (has_range && !autofilter) {
            app->add_part_name(defined_name->app_name);
            named_range_count++;
        }
    }

    /* Add the Named Range heading pairs. */
    if (named_range_count) {
        number = std::to_string(named_range_count);
        app->add_heading_pair("Named Ranges", number);
    }

    /* Set the app/doc properties. */
    app->properties = &workbook->properties;

    app->assemble_xml_file();

    err = _add_file_to_zip(app->file, "docProps/app.xml");
    RETURN_ON_ERROR(err);

    fclose(app->file);

    return 0;
}

/*
 * Write the core.xml file.
 */
uint8_t packager::_write_core_file()
{
    std::shared_ptr<xlsxwriter::core> core = std::make_shared<xlsxwriter::core>();
    int err;

    core->file = lxw_tmpfile(tmpdir.c_str());
    if (!core->file)
        return LXW_ERROR_CREATING_TMPFILE;

    core->properties = &workbook->properties;

    core->assemble_xml_file();

    err = _add_file_to_zip(core->file, "docProps/core.xml");
    RETURN_ON_ERROR(err);

    fclose(core->file);

    return 0;
}

/*
 * Write the custom.xml file.
 */
uint8_t packager::_write_custom_file()
{

    int err;

    if (workbook->custom_properties.empty())
        return 0;

    custom_ptr custom = std::make_shared<xlsxwriter::custom>(workbook->custom_properties);

    custom->file = lxw_tmpfile(tmpdir.c_str());
    if (!custom->file)
        return LXW_ERROR_CREATING_TMPFILE;

    custom->assemble_xml_file();

    err = _add_file_to_zip(custom->file, "docProps/custom.xml");
    RETURN_ON_ERROR(err);

    fclose(custom->file);

    return 0;
}

/*
 * Write the theme.xml file.
 */
uint8_t packager::_write_theme_file()
{
    std::shared_ptr<xlsxwriter::theme> theme = std::make_shared<xlsxwriter::theme>();
    int err;

    theme->file = lxw_tmpfile(tmpdir.c_str());
    if (!theme->file)
        return LXW_ERROR_CREATING_TMPFILE;

    theme->assemble_xml_file();

    err = _add_file_to_zip(theme->file, "xl/theme/theme1.xml");
    RETURN_ON_ERROR(err);

    fclose(theme->file);
    return 0;
}

/*
 * Write the styles.xml file.
 */
uint8_t packager::_write_styles_file()
{
    std::shared_ptr<xlsxwriter::styles> styles = std::make_shared<xlsxwriter::styles>();
    int err;

    /* Copy the unique and in-use formats from the workbook to the styles
     * xf_format list. */
    for (const auto& pair : workbook->used_xf_formats.order_list) {
        /*
        xlsxwriter::format *workbook_format = (xlsxwriter::format *) hash_element->value;
        xlsxwriter::format *style_format = xlsxwriter::format_new();
        memcpy(style_format, workbook_format, sizeof(xlsxwriter::format));
        STAILQ_INSERT_TAIL(styles->xf_formats, style_format, list_pointers);
        */
        styles->xf_formats.push_back(pair.first);
    }

    styles->font_count = workbook->font_count;
    styles->border_count = workbook->border_count;
    styles->fill_count = workbook->fill_count;
    styles->num_format_count = workbook->num_format_count;
    styles->xf_count = workbook->used_xf_formats.order_list.size();

    styles->file = lxw_tmpfile(tmpdir.c_str());
    if (!styles->file)
        return LXW_ERROR_CREATING_TMPFILE;

    styles->assemble_xml_file();

    err = _add_file_to_zip(styles->file, "xl/styles.xml");
    RETURN_ON_ERROR(err);

    fclose(styles->file);

    return 0;
}

/*
 * Write the ContentTypes.xml file.
 */
uint8_t packager::_write_content_types_file()
{
    content_types_ptr content_types = std::make_shared<xlsxwriter::content_types>();
    char filename[LXW_MAX_ATTRIBUTE_LENGTH] = { 0 };
    uint16_t index = 1;
    int err;

    content_types->file = lxw_tmpfile(tmpdir.c_str());
    if (!content_types->file)
        return LXW_ERROR_CREATING_TMPFILE;

    if (workbook->has_png)
        content_types->add_default("png", "image/png");

    if (workbook->has_jpeg)
        content_types->add_default("jpeg", "image/jpeg");

    if (workbook->has_bmp)
        content_types->add_default("bmp", "image/bmp");

    for (const auto& worksheet : workbook->worksheets) {
        (void)worksheet;
        lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                     "/xl/worksheets/sheet%d.xml", index++);
        content_types->add_worksheet_name(filename);
    }

    for (index = 1; index <= chart_count; index++) {
        lxw_snprintf(filename, LXW_FILENAME_LENGTH, "/xl/charts/chart%d.xml",
                     index);
        content_types->add_chart_name(filename);
    }

    for (index = 1; index <= drawing_count; index++) {
        lxw_snprintf(filename, LXW_FILENAME_LENGTH,
                     "/xl/drawings/drawing%d.xml", index);
        content_types->add_drawing_name(filename);
    }

    if (workbook->sst->string_count)
        content_types->add_shared_strings();

    if (!workbook->custom_properties.empty())
        content_types->add_custom_properties();

    content_types->assemble_xml_file();

    err = _add_file_to_zip(content_types->file, "[Content_Types].xml");
    RETURN_ON_ERROR(err);

    fclose(content_types->file);

    return 0;
}

/*
 * Write the workbook .rels xml file.
 */
uint8_t packager::_write_workbook_rels_file()
{
    std::shared_ptr<relationships> rels = std::make_shared<relationships>();
    char sheetname[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;
    int err;

    rels->file = lxw_tmpfile(tmpdir.c_str());
    if (!rels->file)
        return LXW_ERROR_CREATING_TMPFILE;

    for(const auto& worksheet : workbook->worksheets) {
        (void)worksheet;
        lxw_snprintf(sheetname, LXW_FILENAME_LENGTH, "worksheets/sheet%d.xml",
                     index++);
        rels->add_document("/worksheet", sheetname);
    }

    rels->add_document("/theme", "theme/theme1.xml");
    rels->add_document("/styles", "styles.xml");

    if (workbook->sst->string_count)
        rels->add_document("/sharedStrings", "sharedStrings.xml");

    rels->assemble_xml_file();

    err = _add_file_to_zip(rels->file, "xl/_rels/workbook.xml.rels");
    RETURN_ON_ERROR(err);

    fclose(rels->file);

    return 0;
}

/*
 * Write the worksheet .rels files for worksheets that contain links to
 * external data such as hyperlinks or drawings.
 */
uint8_t packager::_write_worksheet_rels_file()
{
    char sheetname[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 0;
    int err;

    for (const auto& worksheet : workbook->worksheets) {

        index++;

        if (worksheet->external_hyperlinks.empty() &&
            worksheet->external_drawing_links.empty())
            continue;

        relationships_ptr rels = std::make_shared<relationships>();
        rels->file = lxw_tmpfile(tmpdir.c_str());
        if (!rels->file)
            return LXW_ERROR_CREATING_TMPFILE;

        for (const auto& rel : worksheet->external_hyperlinks) {
            rels->add_worksheet(rel->type, rel->target, rel->target_mode);
        }

        for (const auto& rel : worksheet->external_drawing_links) {
            rels->add_worksheet(rel->type, rel->target, rel->target_mode);
        }

        lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
                     "xl/worksheets/_rels/sheet%d.xml.rels", index);

        rels->assemble_xml_file();

        err = _add_file_to_zip(rels->file, sheetname);
        RETURN_ON_ERROR(err);

        fclose(rels->file);
    }

    return 0;
}

/*
 * Write the drawing .rels files for worksheets that contain charts or
 * drawings.
 */
uint8_t packager::_write_drawing_rels_file()
{
    char sheetname[LXW_FILENAME_LENGTH] = { 0 };
    uint16_t index = 1;
    int err;

    for (const auto& worksheet : workbook->worksheets) {

        if (worksheet->drawing_links.empty())
            continue;

        std::shared_ptr<relationships> rels = std::make_shared<relationships>();
        rels->file = lxw_tmpfile(tmpdir.c_str());
        if (!rels->file)
            return LXW_ERROR_CREATING_TMPFILE;

        for (const auto& rel : worksheet->drawing_links) {
            rels->add_worksheet(rel->type, rel->target, rel->target_mode);
        }

        lxw_snprintf(sheetname, LXW_FILENAME_LENGTH,
                     "xl/drawings/_rels/drawing%d.xml.rels", index++);

        rels->assemble_xml_file();

        err = _add_file_to_zip(rels->file, sheetname);
        RETURN_ON_ERROR(err);

        fclose(rels->file);
    }

    return 0;
}

/*
 * Write the _rels/.rels xml file.
 */
uint8_t packager::_write_root_rels_file()
{
    relationships_ptr rels = std::make_shared<relationships>();
    int err;

    rels->file = lxw_tmpfile(tmpdir.c_str());
    if (!rels->file)
        return LXW_ERROR_CREATING_TMPFILE;

    rels->add_document("/officeDocument", "xl/workbook.xml");

    rels->add_package("/metadata/core-properties", "docProps/core.xml");

    rels->add_document("/extended-properties", "docProps/app.xml");

    if (!workbook->custom_properties.empty())
        rels->add_document("/custom-properties", "docProps/custom.xml");

    rels->assemble_xml_file();

    err = _add_file_to_zip(rels->file, "_rels/.rels");
    RETURN_ON_ERROR(err);

    fclose(rels->file);

    return 0;
}

/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

uint8_t packager::_add_file_to_zip(FILE * file, const char *filename)
{
    int16_t error = ZIP_OK;
    size_t size_read;
    size_t buffer_size = LXW_ZIP_BUFFER_SIZE;
    char buffer[LXW_ZIP_BUFFER_SIZE];
    memset((void*)buffer, 0, LXW_ZIP_BUFFER_SIZE);

    error = zipOpenNewFileInZip4_64(zipfile,
                                    filename,
                                    &zipfile_info,
                                    NULL, 0, NULL, 0, NULL,
                                    Z_DEFLATED, Z_DEFAULT_COMPRESSION, 0,
                                    -MAX_WBITS, DEF_MEM_LEVEL,
                                    Z_DEFAULT_STRATEGY, NULL, 0, 0, 0, 0);

    if (error != ZIP_OK) {
        LXW_ERROR("Error adding member to zipfile");
        RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
    }

    fflush(file);
    rewind(file);

    size_read = fread(buffer, 1, buffer_size, file);

    while (size_read) {

        if (size_read < buffer_size) {
            if (feof(file) == 0) {
                LXW_ERROR("Error reading member file data");
                RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
            }
        }

        error = zipWriteInFileInZip(zipfile, buffer, (unsigned int) size_read);

        if (error < 0) {
            LXW_ERROR("Error in writing member in the zipfile");
            RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
        }

        size_read = fread(buffer, 1, buffer_size, file);
    }

    if (error < 0) {
        RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
    }
    else {
        error = zipCloseFileInZip(zipfile);
        if (error != ZIP_OK) {
            LXW_ERROR("Error in closing member in the zipfile");
            RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_FILE_ADD);
        }
    }

    return 0;
}

/*
 * Write the xml files that make up the XLXS OPC package.
 */
uint8_t packager::create_package()
{
    int8_t error;

    error = _write_worksheet_files();
    RETURN_ON_ERROR(error);

    error = _write_workbook_file();
    RETURN_ON_ERROR(error);

    error = _write_chart_files();
    RETURN_ON_ERROR(error);

    error = _write_drawing_files();
    RETURN_ON_ERROR(error);

    error = _write_shared_strings_file();
    RETURN_ON_ERROR(error);

    error = _write_app_file();
    RETURN_ON_ERROR(error);

    error = _write_core_file();
    RETURN_ON_ERROR(error);

    error = _write_custom_file();
    RETURN_ON_ERROR(error);

    error = _write_theme_file();
    RETURN_ON_ERROR(error);

    error = _write_styles_file();
    RETURN_ON_ERROR(error);

    error = _write_content_types_file();
    RETURN_ON_ERROR(error);

    error = _write_workbook_rels_file();
    RETURN_ON_ERROR(error);

    error = _write_worksheet_rels_file();
    RETURN_ON_ERROR(error);

    error = _write_drawing_rels_file();
    RETURN_ON_ERROR(error);

    error = _write_image_files();
    RETURN_ON_ERROR(error);;

    error = _write_root_rels_file();
    RETURN_ON_ERROR(error);

    error = zipClose(zipfile, NULL);
    if (error) {
        RETURN_ON_ZIP_ERROR(error, LXW_ERROR_ZIP_CLOSE);
    }

    return 0;
}

} //namespace xlsxwriter
