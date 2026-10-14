//! The decompressed `dir` stream (MS-OVBA section 2.3.4.2): project information, references and
//! the module list. Every record is `Id (u16) Size (u32) Data`, with one documented exception
//! (PROJECTVERSION), so the parser walks records generically and keeps what it needs.

use crate::Error;

/// What the `dir` stream says about one module.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DirModule {
    pub name: String,
    pub stream_name: String,
    /// Offset of the compressed source inside the module stream (MODULEOFFSET).
    pub text_offset: u32,
    /// true for MODULETYPE 0x0022 (document, class or designer module).
    pub non_procedural: bool,
    pub read_only: bool,
    pub private: bool,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DirInfo {
    pub code_page: u16,
    pub project_name: String,
    pub modules: Vec<DirModule>,
}

const PROJECTCODEPAGE: u16 = 0x0003;
const PROJECTNAME: u16 = 0x0004;
const PROJECTVERSION: u16 = 0x0009;
const MODULENAME: u16 = 0x0019;
const MODULESTREAMNAME: u16 = 0x001A;
const MODULETYPE_PROCEDURAL: u16 = 0x0021;
const MODULETYPE_OTHER: u16 = 0x0022;
const MODULEREADONLY: u16 = 0x0025;
const MODULEPRIVATE: u16 = 0x0028;
const MODULETERMINATOR: u16 = 0x002B;
const MODULEOFFSET: u16 = 0x0031;
const MODULESTREAMNAMEUNICODE: u16 = 0x0032;
const MODULENAMEUNICODE: u16 = 0x0047;
const DIR_END: u16 = 0x0010;

pub fn parse(dir: &[u8]) -> Result<DirInfo, Error> {
    let mut info = DirInfo { code_page: 1252, ..DirInfo::default() };
    let mut current: Option<DirModule> = None;
    let mut pos = 0usize;
    // Unicode names win over the code-page ones when both are present.
    let mut name_unicode = false;
    let mut stream_unicode = false;

    while pos + 6 <= dir.len() {
        let id = u16::from_le_bytes([dir[pos], dir[pos + 1]]);
        let size = u32::from_le_bytes([dir[pos + 2], dir[pos + 3], dir[pos + 4], dir[pos + 5]]) as usize;
        pos += 6;
        if id == PROJECTVERSION {
            // Size is the reserved value 4, yet VersionMajor (u32) and VersionMinor (u16) follow.
            pos += 6;
            continue;
        }
        let Some(data) = dir.get(pos..pos.saturating_add(size)) else {
            return Err(Error::Corrupt("dir record runs past the end of the stream"));
        };
        pos += size;

        match id {
            PROJECTCODEPAGE if data.len() >= 2 => info.code_page = u16::from_le_bytes([data[0], data[1]]),
            PROJECTNAME => info.project_name = decode(data, info.code_page),
            MODULENAME => {
                if let Some(done) = current.take() {
                    info.modules.push(done);
                }
                current = Some(DirModule { name: decode(data, info.code_page), ..DirModule::default() });
                name_unicode = false;
                stream_unicode = false;
            }
            MODULENAMEUNICODE => {
                if let Some(m) = current.as_mut() {
                    let n = utf16(data);
                    if !n.is_empty() {
                        m.name = n;
                        name_unicode = true;
                    }
                }
            }
            MODULESTREAMNAME => {
                if let Some(m) = current.as_mut()
                    && !stream_unicode
                {
                    m.stream_name = decode(data, info.code_page);
                }
            }
            MODULESTREAMNAMEUNICODE => {
                if let Some(m) = current.as_mut() {
                    let n = utf16(data);
                    if !n.is_empty() {
                        m.stream_name = n;
                        stream_unicode = true;
                    }
                }
            }
            MODULEOFFSET if data.len() >= 4 => {
                if let Some(m) = current.as_mut() {
                    m.text_offset = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
                }
            }
            MODULETYPE_PROCEDURAL | MODULETYPE_OTHER => {
                if let Some(m) = current.as_mut() {
                    m.non_procedural = id == MODULETYPE_OTHER;
                }
            }
            MODULEREADONLY => {
                if let Some(m) = current.as_mut() {
                    m.read_only = true;
                }
            }
            MODULEPRIVATE => {
                if let Some(m) = current.as_mut() {
                    m.private = true;
                }
            }
            MODULETERMINATOR => {
                if let Some(done) = current.take() {
                    info.modules.push(done);
                }
            }
            DIR_END => break,
            _ => {}
        }
    }
    let _ = name_unicode;
    if let Some(done) = current.take() {
        info.modules.push(done);
    }
    for m in &mut info.modules {
        if m.stream_name.is_empty() {
            m.stream_name = m.name.clone();
        }
    }
    Ok(info)
}

/// Text in the project's code page (MBCS), falling back to Windows-1252.
pub fn decode(bytes: &[u8], code_page: u16) -> String {
    let enc = codepage_encoding(code_page).unwrap_or(encoding_rs::WINDOWS_1252);
    enc.decode(bytes).0.into_owned()
}

fn utf16(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes.as_chunks::<2>().0.iter().map(|p| u16::from_le_bytes(*p)).collect();
    String::from_utf16_lossy(&units)
}

fn codepage_encoding(cp: u16) -> Option<&'static encoding_rs::Encoding> {
    use encoding_rs::*;
    Some(match cp {
        874 => WINDOWS_874,
        932 => SHIFT_JIS,
        936 => GBK,
        949 => EUC_KR,
        950 => BIG5,
        1200 => UTF_16LE,
        1250 => WINDOWS_1250,
        1251 => WINDOWS_1251,
        1252 => WINDOWS_1252,
        1253 => WINDOWS_1253,
        1254 => WINDOWS_1254,
        1255 => WINDOWS_1255,
        1256 => WINDOWS_1256,
        1257 => WINDOWS_1257,
        1258 => WINDOWS_1258,
        10000 => MACINTOSH,
        65001 => UTF_8,
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rec(id: u16, data: &[u8]) -> Vec<u8> {
        let mut v = id.to_le_bytes().to_vec();
        v.extend_from_slice(&(data.len() as u32).to_le_bytes());
        v.extend_from_slice(data);
        v
    }

    #[test]
    fn reads_modules_and_skips_the_version_quirk() {
        let mut dir = Vec::new();
        dir.extend(rec(PROJECTCODEPAGE, &1252u16.to_le_bytes()));
        dir.extend(rec(PROJECTNAME, b"VBAProject"));
        // PROJECTVERSION: Size says 4, six bytes follow.
        dir.extend_from_slice(&PROJECTVERSION.to_le_bytes());
        dir.extend_from_slice(&4u32.to_le_bytes());
        dir.extend_from_slice(&[1, 0, 0, 0, 2, 0]);
        dir.extend(rec(0x000F, &1u16.to_le_bytes()));
        dir.extend(rec(MODULENAME, b"Module1"));
        dir.extend(rec(MODULENAMEUNICODE, &"Module1".encode_utf16().flat_map(u16::to_le_bytes).collect::<Vec<_>>()));
        dir.extend(rec(MODULESTREAMNAME, b"Module1"));
        dir.extend(rec(
            MODULESTREAMNAMEUNICODE,
            &"Module1".encode_utf16().flat_map(u16::to_le_bytes).collect::<Vec<_>>(),
        ));
        dir.extend(rec(MODULEOFFSET, &1234u32.to_le_bytes()));
        dir.extend(rec(MODULETYPE_PROCEDURAL, &[]));
        dir.extend(rec(MODULETERMINATOR, &[]));
        dir.extend(rec(DIR_END, &[]));

        let info = parse(&dir).unwrap();
        assert_eq!(info.project_name, "VBAProject");
        assert_eq!(info.modules.len(), 1);
        assert_eq!(info.modules[0].stream_name, "Module1");
        assert_eq!(info.modules[0].text_offset, 1234);
        assert!(!info.modules[0].non_procedural);
    }

    #[test]
    fn truncated_dir_is_an_error_not_a_panic() {
        let dir = rec(MODULENAME, b"Module1");
        assert!(parse(&dir[..8]).is_err());
    }
}
