/*!
 * @file    ark_zip.c
 * @brief   最小流式 ZIP 写入器实现 —— Store 模式, 两遍法, 零数据缓冲
 */
#include <string.h>
#include "ark_zip.h"

/* ---------------- CRC32 (查表法) ---------------- */
static const uint32_t s_crc_table[256] = {
0x00000000u,0x77073096u,0xEE0E612Cu,0x990951BAu,0x076DC419u,0x706AF48Fu,0xE963A535u,0x9E6495A3u,
0x0EDB8832u,0x79DCB8A4u,0xE0D5E91Eu,0x97D2D988u,0x09B64C2Bu,0x7EB17CBDu,0xE7B82D07u,0x90BF1D91u,
0x1DB71064u,0x6AB020F2u,0xF3B97148u,0x84BE41DEu,0x1ADAD47Du,0x6DDDE4EBu,0xF4D4B551u,0x83D385C7u,
0x136C9856u,0x646BA8C0u,0xFD62F97Au,0x8A65C9ECu,0x14015C4Fu,0x63066CD9u,0xFA0F3D63u,0x8D080DF5u,
0x3B6E20C8u,0x4C69105Eu,0xD56041E4u,0xA2677172u,0x3C03E4D1u,0x4B04D447u,0xD20D85FDu,0xA50AB56Bu,
0x35B5A8FAu,0x42B2986Cu,0xDBBBC9D6u,0xACBCF940u,0x32D86CE3u,0x45DF5C75u,0xDCD60DCFu,0xABD13D59u,
0x26D930ACu,0x51DE003Au,0xC8D75180u,0xBFD06116u,0x21B4F4B5u,0x56B3C423u,0xCFBA9599u,0xB8BDA50Fu,
0x2802B89Eu,0x5F058808u,0xC60CD9B2u,0xB10BE924u,0x2F6F7C87u,0x58684C11u,0xC1611DABu,0xB6662D3Du,
0x76DC4190u,0x01DB7106u,0x98D220BCu,0xEFD5102Au,0x71B18589u,0x06B6B51Fu,0x9FBFE4A5u,0xE8B8D433u,
0x7807C9A2u,0x0F00F934u,0x9609A88Eu,0xE10E9818u,0x7F6A0DBBu,0x086D3D2Du,0x91646C97u,0xE6635C01u,
0x6B6B51F4u,0x1C6C6162u,0x856530D8u,0xF262004Eu,0x6C0695EDu,0x1B01A57Bu,0x8208F4C1u,0xF50FC457u,
0x65B0D9C6u,0x12B7E950u,0x8BBEB8EAu,0xFCB9887Cu,0x62DD1DDFu,0x15DA2D49u,0x8CD37CF3u,0xFBD44C65u,
0x4DB26158u,0x3AB551CEu,0xA3BC0074u,0xD4BB30E2u,0x4ADFA541u,0x3DD895D7u,0xA4D1C46Du,0xD3D6F4FBu,
0x4369E96Au,0x346ED9FCu,0xAD678846u,0xDA60B8D0u,0x44042D73u,0x33031DE5u,0xAA0A4C5Fu,0xDD0D7CC9u,
0x5005713Cu,0x270241AAu,0xBE0B1010u,0xC90C2086u,0x5768B525u,0x206F85B3u,0xB966D409u,0xCE61E49Fu,
0x5EDEF90Eu,0x29D9C998u,0xB0D09822u,0xC7D7A8B4u,0x59B33D17u,0x2EB40D81u,0xB7BD5C3Bu,0xC0BA6CADu,
0xEDB88320u,0x9ABFB3B6u,0x03B6E20Cu,0x74B1D29Au,0xEAD54739u,0x9DD277AFu,0x04DB2615u,0x73DC1683u,
0xE3630B12u,0x94643B84u,0x0D6D6A3Eu,0x7A6A5AA8u,0xE40ECF0Bu,0x9309FF9Du,0x0A00AE27u,0x7D079EB1u,
0xF00F9344u,0x8708A3D2u,0x1E01F268u,0x6906C2FEu,0xF762575Du,0x806567CBu,0x196C3671u,0x6E6B06E7u,
0xFED41B76u,0x89D32BE0u,0x10DA7A5Au,0x67DD4ACCu,0xF9B9DF6Fu,0x8EBEEFF9u,0x17B7BE43u,0x60B08ED5u,
0xD6D6A3E8u,0xA1D1937Eu,0x38D8C2C4u,0x4FDFF252u,0xD1BB67F1u,0xA6BC5767u,0x3FB506DDu,0x48B2364Bu,
0xD80D2BDAu,0xAF0A1B4Cu,0x36034AF6u,0x41047A60u,0xDF60EFC3u,0xA867DF55u,0x316E8EEFu,0x4669BE79u,
0xCB61B38Cu,0xBC66831Au,0x256FD2A0u,0x5268E236u,0xCC0C7795u,0xBB0B4703u,0x220216B9u,0x5505262Fu,
0xC5BA3BBEu,0xB2BD0B28u,0x2BB45A92u,0x5CB36A04u,0xC2D7FFA7u,0xB5D0CF31u,0x2CD99E8Bu,0x5BDEAE1Du,
0x9B64C2B0u,0xEC63F226u,0x756AA39Cu,0x026D930Au,0x9C0906A9u,0xEB0E363Fu,0x72076785u,0x05005713u,
0x95BF4A82u,0xE2B87A14u,0x7BB12BAEu,0x0CB61B38u,0x92D28E9Bu,0xE5D5BE0Du,0x7CDCEFB7u,0x0BDBDF21u,
0x86D3D2D4u,0xF1D4E242u,0x68DDB3F8u,0x1FDA836Eu,0x81BE16CDu,0xF6B9265Bu,0x6FB077E1u,0x18B74777u,
0x88085AE6u,0xFF0F6A70u,0x66063BCAu,0x11010B5Cu,0x8F659EFFu,0xF862AE69u,0x616BFFD3u,0x166CCF45u,
0xA00AE278u,0xD70DD2EEu,0x4E048354u,0x3903B3C2u,0xA7672661u,0xD06016F7u,0x4969474Du,0x3E6E77DBu,
0xAED16A4Au,0xD9D65ADCu,0x40DF0B66u,0x37D83BF0u,0xA9BCAE53u,0xDEBB9EC5u,0x47B2CF7Fu,0x30B5FFE9u,
0xBDBDF21Cu,0xCABAC28Au,0x53B39330u,0x24B4A3A6u,0xBAD03605u,0xCDD70693u,0x54DE5729u,0x23D967BFu,
0xB3667A2Eu,0xC4614AB8u,0x5D681B02u,0x2A6F2B94u,0xB40BBE37u,0xC30C8EA1u,0x5A05DF1Bu,0x2D02EF8Du
};

uint32_t ark_crc32(uint32_t crc, const void *data, size_t len)
{
    const uint8_t *p = (const uint8_t *)data;
    crc = ~crc;
    while (len--)
        crc = s_crc_table[(crc ^ *p++) & 0xFFu] ^ (crc >> 8);
    return ~crc;
}

/* ---------------- 小工具 ---------------- */
static void put16(uint8_t *p, uint16_t v)
{
    p[0] = (uint8_t)(v & 0xFF);
    p[1] = (uint8_t)(v >> 8);
}
static void put32(uint8_t *p, uint32_t v)
{
    p[0] = (uint8_t)(v & 0xFF);
    p[1] = (uint8_t)((v >> 8) & 0xFF);
    p[2] = (uint8_t)((v >> 16) & 0xFF);
    p[3] = (uint8_t)((v >> 24) & 0xFF);
}

void ark_zip_begin(ark_zip_t *z, ark_zip_write_fn write, void *ctx)
{
    memset(z, 0, sizeof(*z));
    z->write = write;
    z->ctx   = ctx;
    z->cur   = 0xFFFFu;
}

/* ---------------- 第一遍: 测量 ---------------- */
void ark_zip_measure_start(ark_zip_t *z)
{
    z->m_crc  = 0;
    z->m_size = 0;
}
void ark_zip_measure_write(ark_zip_t *z, const void *data, size_t len)
{
    z->m_crc  = ark_crc32(z->m_crc, data, len);
    z->m_size += (uint32_t)len;
}

/* ---------------- 第二遍: 真实输出 ---------------- */
static int emit(ark_zip_t *z, const void *data, size_t len)
{
    if (z->out_failed) return -1;
    if (z->write && z->write(data, len, z->ctx) != 0) {
        z->out_failed = 1;
        return -1;
    }
    z->out_pos += (uint32_t)len;
    return 0;
}

int ark_zip_open_stream(ark_zip_t *z, const char *name)
{
    return ark_zip_open_raw(z, name, 0, z->m_crc, z->m_size, z->m_size);
}

int ark_zip_open_raw(ark_zip_t *z, const char *name,
                     uint16_t method, uint32_t crc,
                     uint32_t usize, uint32_t csize)
{
    uint8_t hdr[30];
    size_t n;
    ark_zip_entry_t *e;
    if (z->entry_count >= ARK_ZIP_MAX_ENTRIES) return -1;
    n = strlen(name);
    if (n == 0 || n >= sizeof(e->name)) return -2;

    z->cur = z->entry_count++;
    e = &z->entries[z->cur];
    memset(e, 0, sizeof(*e));
    memcpy(e->name, name, n);
    e->size   = csize;             /* Store: 大小一致; 记录用 csize */
    e->usize  = usize;
    e->crc    = crc;
    e->method = method;
    e->offset = z->out_pos;
    /* 固定时间戳 2024-01-01 00:00:00 (DOS 格式), 不依赖 RTC */
    e->dosdate = (uint16_t)(((2024 - 1980) << 9) | (1 << 5) | 1);
    e->dostime = 0;

    memset(hdr, 0, sizeof(hdr));
    put32(hdr + 0,  0x04034B50u);   /* local file header 签名 */
    put16(hdr + 4,  20);            /* version needed */
    put16(hdr + 6,  0x0800u);       /* 通用标志: UTF-8 文件名 */
    put16(hdr + 8,  method);
    put16(hdr + 10, e->dostime);
    put16(hdr + 12, e->dosdate);
    put32(hdr + 14, e->crc);
    put32(hdr + 18, csize);        /* 压缩后大小 */
    put32(hdr + 22, usize);        /* 原始大小 */
    put16(hdr + 26, (uint16_t)n);
    put16(hdr + 28, 0);
    if (emit(z, hdr, sizeof(hdr))) return -3;
    if (emit(z, name, n)) return -3;
    return 0;
}

int ark_zip_write(ark_zip_t *z, const void *data, size_t len)
{
    if (z->cur == 0xFFFFu) return -1;
    return emit(z, data, len);
}

int ark_zip_close_stream(ark_zip_t *z)
{
    if (z->cur == 0xFFFFu) return -1;
    z->cur = 0xFFFFu;
    return z->out_failed ? -2 : 0;
}

int ark_zip_end(ark_zip_t *z)
{
    uint8_t cd[46], eocd[22];
    uint32_t cd_start, cd_size;
    uint16_t i;
    if (z->cur != 0xFFFFu) return -1;

    cd_start = z->out_pos;
    for (i = 0; i < z->entry_count; i++) {
        ark_zip_entry_t *e = &z->entries[i];
        size_t n = strlen(e->name);
        memset(cd, 0, sizeof(cd));
        put32(cd + 0,  0x02014B50u);
        put16(cd + 4,  20);
        put16(cd + 6,  20);
        put16(cd + 8,  0x0800u);
        put16(cd + 10, e->method);
        put16(cd + 12, e->dostime);
        put16(cd + 14, e->dosdate);
        put32(cd + 16, e->crc);
        put32(cd + 20, e->size);    /* 压缩后 */
        put32(cd + 24, e->usize);   /* 原始 */
        put16(cd + 28, (uint16_t)n);
        put32(cd + 42, e->offset);
        if (emit(z, cd, sizeof(cd))) return -2;
        if (emit(z, e->name, n)) return -2;
    }
    cd_size = z->out_pos - cd_start;

    memset(eocd, 0, sizeof(eocd));
    put32(eocd + 0,  0x06054B50u);
    put16(eocd + 8,  z->entry_count);
    put16(eocd + 10, z->entry_count);
    put32(eocd + 12, cd_size);
    put32(eocd + 16, cd_start);
    if (emit(z, eocd, sizeof(eocd))) return -2;
    return z->out_failed ? -2 : 0;
}
