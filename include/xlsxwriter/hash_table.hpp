/*
 * libxlsxwriter
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * hash_table - Hash table functions for libxlsxwriter.
 *
 */

#ifndef __LXW_HASH_TABLE_H__
#define __LXW_HASH_TABLE_H__

#include "common.hpp"
#include <list>
#include <stdint.h>
#include <unordered_map>

/* Macro to loop over hash table elements in insertion order. */
#define LXW_FOREACH_ORDERED(elem, hash_table) \
    STAILQ_FOREACH((elem), (hash_table)->order_list, lxw_hash_order_pointers)

/* List declarations. */
STAILQ_HEAD(lxw_hash_order_list, lxw_hash_element);
SLIST_HEAD(lxw_hash_bucket_list, lxw_hash_element);

/* LXW_HASH hash table struct. */
typedef struct lxw_hash_table {
    uint32_t num_buckets;
    uint32_t used_buckets;
    uint32_t unique_count;
    uint8_t free_key;
    uint8_t free_value;

    struct lxw_hash_order_list *order_list;
    struct lxw_hash_bucket_list **buckets;
} lxw_hash_table;

/*
 * LXW_HASH table element struct.
 *
 * The hash elements contain pointers to allow them to be stored in
 * lists in the the hash table buckets and also pointers to track the
 * insertion order in a separate list.
 */
typedef struct lxw_hash_element {
    void *key;
    void *value;

    STAILQ_ENTRY (lxw_hash_element) lxw_hash_order_pointers;
    SLIST_ENTRY (lxw_hash_element) lxw_hash_list_pointers;
} lxw_hash_element;


size_t _generate_hash_key(void *data, size_t data_len, size_t num_buckets);

template <class T>
size_t generate_hash_key(const T* data)
{
    return _generate_hash_key((void*)data, sizeof(T), INT16_MAX);
}
size_t generate_hash_key(const std::string& data);


lxw_hash_element *lxw_hash_key_exists(lxw_hash_table *lxw_hash, void *key,
                                      size_t key_len);
lxw_hash_element *lxw_insert_hash_element(lxw_hash_table *lxw_hash, void *key,
                                          void *value, size_t key_len);
lxw_hash_table *lxw_hash_new(uint32_t num_buckets, uint8_t free_key,
                             uint8_t free_value);
void lxw_hash_free(lxw_hash_table *lxw_hash);

namespace xlsxwriter {

template <class K, class V>
class hash_table {
public:

    std::pair<std::pair<K, V>, bool > exists(const K& key) {
        size_t keyhash = generate_hash_key(key);
        auto it = storage.find(keyhash);
        if (it != storage.end())
            return std::make_pair(it->second, true);
        else
            return std::make_pair(std::pair<K, V>(), false);
    }

    std::pair<std::pair<K, V>, bool > insert(const K& key, const V& val) {
        size_t keyhash = generate_hash_key(key);
        auto res = storage.insert(std::make_pair(keyhash, std::make_pair(key, val)));
        if (res.second)
            order_list.push_back(std::make_pair(key, val));
        return std::make_pair(res.first->second, res.second);
    }

    std::list<std::pair<K, V>> order_list;
    std::unordered_map<size_t, std::pair<K, V>> storage;
};

}

#endif /* __LXW_HASH_TABLE_H__ */
