//! Password verifiers of `sheetProtection` and `workbookProtection` (ECMA-376 Part 1, 18.3.1.85
//! and 18.2.29): the legacy hex hash attribute or the `algorithmName` / `hashValue` /
//! `saltValue` / `spinCount` set (prefixed `workbook` on the workbook element).

use macroost_sheet::PasswordHash;
use macroost_sheet::protect::{base64_decode, base64_encode};

/// Reads a verifier through `get` (attribute lookup). `legacy` names the hex attribute
/// (`password` / `workbookPassword`); `prefix` is `""` or `"workbook"`.
pub(crate) fn read_hash(get: &dyn Fn(&str) -> Option<String>, legacy: &str, prefix: &str) -> Option<PasswordHash> {
    let name = |n: &str| {
        if prefix.is_empty() {
            n.to_owned()
        } else {
            let mut c = n.chars();
            let first = c.next().map(|f| f.to_ascii_uppercase().to_string()).unwrap_or_default();
            format!("{prefix}{first}{}", c.as_str())
        }
    };
    if let (Some(alg), Some(hash), Some(salt)) =
        (get(&name("algorithmName")), get(&name("hashValue")), get(&name("saltValue")))
    {
        let spin = get(&name("spinCount")).and_then(|s| s.trim().parse().ok()).unwrap_or(0);
        return Some(PasswordHash::Modern {
            algorithm: alg,
            hash: base64_decode(&hash)?,
            salt: base64_decode(&salt)?,
            spin_count: spin,
        });
    }
    let hex = get(legacy)?;
    u16::from_str_radix(hex.trim(), 16).ok().filter(|h| *h != 0).map(PasswordHash::Legacy)
}

/// Attribute text for a verifier (leading space included), or empty for none.
pub(crate) fn hash_attrs(h: Option<&PasswordHash>, legacy: &str, prefix: &str) -> String {
    let name = |n: &str| {
        if prefix.is_empty() {
            n.to_owned()
        } else {
            let mut c = n.chars();
            let first = c.next().map(|f| f.to_ascii_uppercase().to_string()).unwrap_or_default();
            format!("{prefix}{first}{}", c.as_str())
        }
    };
    match h {
        None => String::new(),
        Some(PasswordHash::Legacy(v)) => format!(" {legacy}=\"{v:04X}\""),
        Some(PasswordHash::Modern { algorithm, hash, salt, spin_count }) => format!(
            " {}=\"{}\" {}=\"{}\" {}=\"{}\" {}=\"{spin_count}\"",
            name("algorithmName"),
            algorithm,
            name("hashValue"),
            base64_encode(hash),
            name("saltValue"),
            base64_encode(salt),
            name("spinCount"),
        ),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip() {
        let h = PasswordHash::Modern {
            algorithm: "SHA-512".into(),
            hash: vec![1, 2, 3],
            salt: vec![9, 8],
            spin_count: 100,
        };
        let text = hash_attrs(Some(&h), "workbookPassword", "workbook");
        let get = |n: &str| -> Option<String> {
            let key = format!(" {n}=\"");
            let i = text.find(&key)? + key.len();
            Some(text[i..].split('"').next()?.to_owned())
        };
        assert_eq!(read_hash(&get, "workbookPassword", "workbook"), Some(h));
        let l = PasswordHash::Legacy(0xCC1A);
        let text = hash_attrs(Some(&l), "password", "");
        assert_eq!(text, " password=\"CC1A\"");
    }
}
