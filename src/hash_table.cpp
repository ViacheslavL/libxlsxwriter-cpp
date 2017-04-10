/*****************************************************************************
 * hash_table - Hash table functions for libxlsxwriter.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <stdint.h>
#include "xlsxwriter/hash_table.hpp"

/*
 * Calculate the hash key using the FNV function. See:
 * http://en.wikipedia.org/wiki/Fowler-Noll-Vo_hash_function
 */
size_t _generate_hash_key(void *data, size_t data_len, size_t num_buckets)
{
    unsigned char *p = (unsigned char*)data;
    size_t hash = 2166136261U;
    size_t i;

    for (i = 0; i < data_len; i++)
        hash = (hash * 16777619) ^ p[i];

    return hash % num_buckets;
}
size_t generate_hash_key(const std::string &data)
{
    return _generate_hash_key((void*)data.c_str(), data.size(), INT16_MAX);
}
