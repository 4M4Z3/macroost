//! The MS-OVBA compression format (section 2.4.1): a signature byte followed by chunks of at most
//! 4096 decompressed bytes, each either raw or a sequence of literal bytes and copy tokens.

use crate::Error;

const CHUNK: usize = 4096;

/// BitCount, LengthMask and the maximum copy length for a copy token written when `pos` bytes of
/// the current chunk have been decompressed (section 2.4.1.3.19.1).
fn copy_token_help(pos: usize) -> (u32, u16, usize) {
    // Smallest bit count with 2^bits >= pos, at least 4.
    let mut bits = 4u32;
    while (1usize << bits) < pos {
        bits += 1;
    }
    let length_mask = 0xFFFFu16 >> bits;
    (bits, length_mask, usize::from(length_mask) + 3)
}

/// Decompresses a CompressedContainer (section 2.4.1.3.1).
pub fn decompress(data: &[u8]) -> Result<Vec<u8>, Error> {
    let Some((&sig, mut rest)) = data.split_first() else {
        return Err(Error::Corrupt("empty compressed container"));
    };
    if sig != 0x01 {
        return Err(Error::Corrupt("compressed container signature is not 0x01"));
    }
    let mut out = Vec::with_capacity(data.len() * 3);
    while rest.len() >= 2 {
        let header = u16::from_le_bytes([rest[0], rest[1]]);
        let size = usize::from(header & 0x0FFF) + 3;
        let compressed = header & 0x8000 != 0;
        let end = size.min(rest.len());
        let body = &rest[2..end];
        let chunk_start = out.len();
        if compressed {
            decompress_chunk(body, &mut out, chunk_start)?;
        } else {
            // A raw chunk holds 4096 bytes; a truncated file keeps whatever is there.
            out.extend_from_slice(&body[..body.len().min(CHUNK)]);
        }
        rest = &rest[end..];
    }
    Ok(out)
}

fn decompress_chunk(mut body: &[u8], out: &mut Vec<u8>, chunk_start: usize) -> Result<(), Error> {
    while let Some((&flags, tail)) = body.split_first() {
        body = tail;
        for bit in 0..8 {
            if body.is_empty() {
                break;
            }
            if flags >> bit & 1 == 0 {
                out.push(body[0]);
                body = &body[1..];
            } else {
                let Some(pair) = body.get(..2) else {
                    return Err(Error::Corrupt("copy token cut short"));
                };
                let token = u16::from_le_bytes([pair[0], pair[1]]);
                body = &body[2..];
                let (bits, length_mask, _) = copy_token_help(out.len() - chunk_start);
                let length = usize::from(token & length_mask) + 3;
                let offset = usize::from((token & !length_mask) >> (16 - bits)) + 1;
                if offset > out.len() - chunk_start {
                    return Err(Error::Corrupt("copy token reaches before the chunk"));
                }
                // Byte by byte: a copy may overlap the bytes it is producing.
                let from = out.len() - offset;
                for i in 0..length {
                    let b = out[from + i];
                    out.push(b);
                }
            }
        }
    }
    Ok(())
}

/// Compresses `data` into a CompressedContainer (section 2.4.1.3.7). Used to write source-only
/// projects and to test the decompressor.
pub fn compress(data: &[u8]) -> Vec<u8> {
    let mut out = vec![0x01];
    for chunk in data.chunks(CHUNK) {
        let body = compress_chunk(chunk);
        if body.len() > CHUNK && chunk.len() == CHUNK {
            // Raw chunk: flag 0, size field 4096 - 3 + 2 (section 2.4.1.3.8).
            out.extend_from_slice(&(0x3000u16 | 0x0FFF).to_le_bytes());
            out.extend_from_slice(chunk);
        } else {
            let header = 0xB000u16 | ((body.len() + 2 - 3) as u16 & 0x0FFF);
            out.extend_from_slice(&header.to_le_bytes());
            out.extend_from_slice(&body);
        }
    }
    out
}

fn compress_chunk(chunk: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(chunk.len() + chunk.len() / 8 + 1);
    let mut pos = 0;
    while pos < chunk.len() {
        let flag_at = out.len();
        out.push(0);
        let mut flags = 0u8;
        for bit in 0..8 {
            if pos >= chunk.len() {
                break;
            }
            let (offset, length) = longest_match(chunk, pos);
            if length >= 3 {
                let (bits, _, _) = copy_token_help(pos);
                let token = (((offset - 1) as u16) << (16 - bits)) | (length - 3) as u16;
                out.extend_from_slice(&token.to_le_bytes());
                flags |= 1 << bit;
                pos += length;
            } else {
                out.push(chunk[pos]);
                pos += 1;
            }
        }
        out[flag_at] = flags;
    }
    out
}

/// The longest earlier match for the bytes at `pos` (section 2.4.1.3.19.4).
fn longest_match(chunk: &[u8], pos: usize) -> (usize, usize) {
    let (_, _, max_len) = copy_token_help(pos);
    let (mut best_off, mut best_len) = (0, 0);
    for start in (0..pos).rev() {
        let mut len = 0;
        while pos + len < chunk.len() && len < max_len && chunk[start + len] == chunk[pos + len] {
            len += 1;
        }
        if len > best_len {
            best_len = len;
            best_off = pos - start;
        }
    }
    (best_off, best_len)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn hex(s: &str) -> Vec<u8> {
        s.split_whitespace().filter_map(|h| u8::from_str_radix(h, 16).ok()).collect()
    }

    /// The three worked examples of section 3.2, byte for byte in both directions.
    #[test]
    fn spec_examples_match_exactly() {
        let cases: [(&[u8], &str); 3] = [
            (
                b"abcdefghijklmnopqrstuv.",
                "01 19 B0 00 61 62 63 64 65 66 67 68 00 69 6A 6B 6C 6D 6E 6F 70 00 71 72 73 74 75 76 2E",
            ),
            (
                b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa",
                "01 2F B0 00 23 61 61 61 62 63 64 65 82 66 00 70 61 67 68 69 6A 01 38 08 61 6B 6C 00 20 6D 6E 6F                  70 06 71 02 70 04 00 72 73 74 75 76 10 77 78 79 7A 00 2C",
            ),
            (&[b'a'; 73], "01 03 B0 02 61 45 00"),
        ];
        for (plain, packed) in cases {
            let packed = hex(packed);
            assert_eq!(decompress(&packed).unwrap(), plain, "decompress");
            assert_eq!(compress(plain), packed, "compress");
        }
    }

    /// Larger inputs, including several chunks.
    #[test]
    fn round_trips_repetitive_and_random_data() {
        let mut seed = 0x2545F491u32;
        let mut random = Vec::new();
        for _ in 0..20_000 {
            seed ^= seed << 13;
            seed ^= seed >> 17;
            seed ^= seed << 5;
            random.push((seed % 7) as u8 + b'a');
        }
        for input in [vec![b'a'; 9000], random, Vec::new(), b"Attribute VB_Name = \"Module1\"\r\n".to_vec()] {
            assert_eq!(decompress(&compress(&input)).unwrap(), input);
        }
    }

    #[test]
    fn rejects_garbage_without_panicking() {
        for bad in [&[][..], &[0x02, 0, 0], &[0x01, 0xFF, 0xBF, 0xFF, 0xFF, 0xFF]] {
            let _ = decompress(bad);
        }
    }
}
