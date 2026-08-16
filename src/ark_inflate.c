/*!
 * @file    ark_inflate.c
 * @brief   自包含 DEFLATE 解压实现(固定/动态 Huffman + Stored 块)
 *          参考 Mark Adler 的 puff 算法结构, 改为拉取式输入。
 */
#include <string.h>
#include "ark_inflate.h"

#define MAXBITS   15
#define MAXLCODES 286
#define MAXDCODES 30
#define MAXCODES  (MAXLCODES + MAXDCODES)
#define FIXLCODES 288

typedef struct {
    ark_inflate_pull pull;
    void            *ctx;
    uint8_t          inbuf[1024];
    uint32_t         inlen, inpos;

    uint8_t         *out;
    size_t           outcap, outcnt;

    int              bitbuf;
    int              bitcnt;
} inf_t;

static int inf_next(inf_t *s)
{
    if (s->inpos >= s->inlen) {
        uint32_t n = s->pull(s->ctx, s->inbuf, sizeof(s->inbuf));
        if (n == 0) return -1;
        s->inlen = n;
        s->inpos = 0;
    }
    return s->inbuf[s->inpos++];
}

static int inf_bits(inf_t *s, int need)
{
    long val = s->bitbuf;
    while (s->bitcnt < need) {
        int c = inf_next(s);
        if (c < 0) return -1;
        val |= (long)c << s->bitcnt;
        s->bitcnt += 8;
    }
    s->bitbuf = (int)(val >> need);
    s->bitcnt -= need;
    return (int)(val & ((1L << need) - 1));
}

static int inf_stored(inf_t *s)
{
    unsigned len, n;
    /* 丢弃未对齐的比特, 读取 LEN/NLEN */
    s->bitbuf = 0;
    s->bitcnt = 0;
    {
        int b0 = inf_bits(s, 8), b1 = inf_bits(s, 8);
        int b2 = inf_bits(s, 8), b3 = inf_bits(s, 8);
        if (b0 < 0 || b1 < 0 || b2 < 0 || b3 < 0) return -1;
        len  = (unsigned)b0 | ((unsigned)b1 << 8);
        n    = (unsigned)b2 | ((unsigned)b3 << 8);
    }
    if (n != (~len & 0xFFFFu)) return -1;
    if (s->outcnt + len > s->outcap) return -2;
    while (len--) {
        int c = inf_next(s);
        if (c < 0) return -1;
        s->out[s->outcnt++] = (uint8_t)c;
    }
    return 0;
}

typedef struct {
    short count[MAXBITS + 1];
    short symbol[MAXCODES];
} huff_t;

static int inf_decode(inf_t *s, const huff_t *h)
{
    int code = 0, first = 0, index = 0, len, count;
    for (len = 1; len <= MAXBITS; len++) {
        int b = inf_bits(s, 1);
        if (b < 0) return -1;
        code |= b;
        count = h->count[len];
        if (code - count < first)
            return h->symbol[index + (code - first)];
        index += count;
        first += count;
        first <<= 1;
        code <<= 1;
    }
    return -1;
}

static int inf_build(huff_t *h, const short *length, int n)
{
    int symbol, len, left;
    short offs[MAXBITS + 1];

    for (len = 0; len <= MAXBITS; len++)
        h->count[len] = 0;
    for (symbol = 0; symbol < n; symbol++)
        h->count[length[symbol]]++;
    if (h->count[0] == n) return 0;

    left = 1;
    for (len = 1; len <= MAXBITS; len++) {
        left <<= 1;
        left -= h->count[len];
        if (left < 0) return left;
    }

    offs[1] = 0;
    for (len = 1; len < MAXBITS; len++)
        offs[len + 1] = offs[len] + h->count[len];
    for (symbol = 0; symbol < n; symbol++)
        if (length[symbol] != 0)
            h->symbol[offs[length[symbol]]++] = (short)symbol;
    return left;
}

static const short LENS[29] = {
    3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,
    195,227,258 };
static const short LEXT[29] = {
    0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0 };
static const short DISTS[30] = {
    1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,
    3073,4097,6145,8193,12289,16385,24577 };
static const short DEXT[30] = {
    0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13 };

static int inf_codes(inf_t *s, const huff_t *lencode, const huff_t *distcode)
{
    int symbol, len, dist;
    do {
        symbol = inf_decode(s, lencode);
        if (symbol < 0) return -1;
        if (symbol < 256) {
            if (s->outcnt + 1 > s->outcap) return -2;
            s->out[s->outcnt++] = (uint8_t)symbol;
        } else if (symbol > 256) {
            symbol -= 257;
            if (symbol >= 29) return -1;
            {
                int extra = inf_bits(s, LEXT[symbol]);
                if (extra < 0) return -1;
                len = LENS[symbol] + extra;
            }
            symbol = inf_decode(s, distcode);
            if (symbol < 0 || symbol >= 30) return -1;
            {
                int extra = inf_bits(s, DEXT[symbol]);
                if (extra < 0) return -1;
                dist = DISTS[symbol] + extra;
            }
            if (s->outcnt + (size_t)len > s->outcap) return -2;
            if ((size_t)dist > s->outcnt) return -1;
            while (len--) {
                s->out[s->outcnt] = s->out[s->outcnt - dist];
                s->outcnt++;
            }
        }
    } while (symbol != 256);
    return 0;
}

static int inf_fixed(inf_t *s)
{
    static int built = 0;
    static huff_t lencode, distcode;
    if (!built) {
        short lengths[FIXLCODES];
        int symbol;
        for (symbol = 0; symbol < 144; symbol++) lengths[symbol] = 8;
        for (; symbol < 256; symbol++)      lengths[symbol] = 9;
        for (; symbol < 280; symbol++)      lengths[symbol] = 7;
        for (; symbol < FIXLCODES; symbol++)lengths[symbol] = 8;
        if (inf_build(&lencode, lengths, FIXLCODES) < 0) return -1;
        for (symbol = 0; symbol < 30; symbol++) lengths[symbol] = 5;
        /* 固定距离表 30个5位码是不完全码(32槽位用30), 仅拒绝负错误 */
        if (inf_build(&distcode, lengths, 30) < 0) return -1;
        built = 1;
    }
    return inf_codes(s, &lencode, &distcode);
}

static int inf_dynamic(inf_t *s)
{
    static const short ORDER[19] = {
        16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15 };
    int nlen, ndist, ncode, index;
    short lengths[MAXCODES];
    huff_t lencode, distcode;
    int symbol, err;

    nlen  = inf_bits(s, 5); if (nlen  < 0) return -1; nlen  += 257;
    ndist = inf_bits(s, 5); if (ndist < 0) return -1; ndist += 1;
    ncode = inf_bits(s, 4); if (ncode < 0) return -1; ncode += 4;
    if (nlen > MAXLCODES || ndist > MAXDCODES) return -1;

    for (index = 0; index < ncode; index++) {
        int b = inf_bits(s, 3);
        if (b < 0) return -1;
        lengths[ORDER[index]] = (short)b;
    }
    for (; index < 19; index++)
        lengths[ORDER[index]] = 0;

    err = inf_build(&lencode, lengths, 19);
    if (err) return -1;

    index = 0;
    while (index < nlen + ndist) {
        symbol = inf_decode(s, &lencode);
        if (symbol < 0) return -1;
        if (symbol < 16) {
            lengths[index++] = (short)symbol;
        } else {
            int len = 0, rep;
            if (symbol == 16) {
                if (index == 0) return -1;
                len = lengths[index - 1];
                symbol = inf_bits(s, 2);
                if (symbol < 0) return -1;
                rep = 3 + symbol;
            } else if (symbol == 17) {
                symbol = inf_bits(s, 3);
                if (symbol < 0) return -1;
                rep = 3 + symbol;
            } else {
                symbol = inf_bits(s, 7);
                if (symbol < 0) return -1;
                rep = 11 + symbol;
            }
            if (index + rep > nlen + ndist) return -1;
            while (rep--)
                lengths[index++] = (short)len;
        }
    }
    if (lengths[256] == 0) return -1;
    err = inf_build(&lencode, lengths, nlen);
    if (err && (err < 0 || nlen != lencode.count[0] + lencode.count[1]))
        return -1;
    err = inf_build(&distcode, lengths + nlen, ndist);
    if (err && (err < 0 || ndist != distcode.count[0] + distcode.count[1]))
        return -1;

    return inf_codes(s, &lencode, &distcode);
}

int ark_inflate(ark_inflate_pull pull, void *ctx,
                uint8_t *out, size_t outcap, size_t *outlen)
{
    inf_t s;
    int last, err = 0;
    memset(&s, 0, sizeof(s));
    s.pull = pull; s.ctx = ctx;
    s.out = out; s.outcap = outcap; s.outcnt = 0;

    do {
        int type;
        last = inf_bits(&s, 1);
        if (last < 0) { err = -1; break; }
        type = inf_bits(&s, 2);
        if (type < 0) { err = -1; break; }
        if      (type == 0) err = inf_stored(&s);
        else if (type == 1) err = inf_fixed(&s);
        else if (type == 2) err = inf_dynamic(&s);
        else                err = -1;
    } while (!err && !last);

    if (err == 0 && outlen) *outlen = s.outcnt;
    return err;
}
