// Forked from https://github.com/pdf-rs/pdf/blob/master/pdf/examples/text.rs
use pdf::content::*;
use pdf::encoding::BaseEncoding;
use pdf::error::PdfError;
use pdf::file::FileOptions;
use pdf::font::*;
use pdf::object::{MaybeRef, PlainRef, RcRef, Resolve};
use std::collections::HashMap;
use std::convert::TryInto;

struct FontInfo {
    font: RcRef<Font>,
    cmap: ToUnicodeMap,
}
struct Cache {
    fonts: HashMap<String, FontInfo>,
}
impl Cache {
    fn new() -> Self {
        Cache {
            fonts: HashMap::new(),
        }
    }
    fn add_font(&mut self, name: impl Into<String>, font: RcRef<Font>, resolver: &impl Resolve) {
        if let Some(to_unicode) = font.to_unicode(resolver) {
            self.fonts.insert(
                name.into(),
                FontInfo {
                    font,
                    cmap: to_unicode.unwrap(),
                },
            );
        }
    }
    fn get_font(&self, name: &str) -> Option<&FontInfo> {
        self.fonts.get(name)
    }
}
fn add_string(data: &[u8], out: &mut String, info: &FontInfo) {
    if let Some(encoding) = info.font.encoding() {
        match encoding.base {
            BaseEncoding::IdentityH => {
                for w in data.windows(2) {
                    let cp = u16::from_be_bytes(w.try_into().unwrap());
                    if let Some(s) = info.cmap.get(cp) {
                        out.push_str(s);
                    }
                }
            }
            _ => {
                for &b in data {
                    if let Some(s) = info.cmap.get(b as u16) {
                        out.push_str(s);
                    } else {
                        out.push(b as char);
                    }
                }
            }
        };
    }
}
fn read_text() -> Result<(), PdfError> {
    let path = r"0.pdf";
    println!("读取PDF文件: {}", path);

    let file = FileOptions::cached().open(&path).unwrap();
    let resolver = file.resolver();

    let mut out = String::new();
    for page in file.pages() {
        let page = page?;
        let resources = page.resources.as_ref().unwrap();
        let mut cache = Cache::new();

        // make sure all fonts are in the cache, so we can reference them
        for (name, font) in resources.fonts.clone() {
            match font {
                MaybeRef::Indirect(font) => {
                    cache.add_font(name.as_str(), font.clone(), &resolver);
                },
                _ => {}
            }
        }
        for gs in resources.graphics_states.values() {
            if let Some((font, _)) = gs.font {
                let font = resolver.get(font)?;
                if let Some(font_name) = &font.name {
                    cache.add_font(font_name.as_str(), font.clone(), &resolver);
                }
            }
        }
        let mut current_font = None;
        let contents = page.contents.as_ref().unwrap();
        for op in contents.operations(&resolver)?.iter() {
            match op {
                Op::GraphicsState { name } => {
                    let gs = resources.graphics_states.get(name).unwrap();

                    if let Some((font_ref, _)) = gs.font {
                        let font = resolver.get(font_ref)?;
                        if let Some(font_name) = &font.name {
                            current_font = cache.get_font(font_name.as_str());
                        }
                    }
                }
                // text font
                Op::TextFont { name, .. } => {
                    current_font = cache.get_font(name.as_str());
                }
                Op::TextDraw { text } => {
                    if let Some(font) = current_font {
                        add_string(&text.data, &mut out, font);
                    }
                }
                Op::TextDrawAdjusted { array } => {
                    if let Some(font) = current_font {
                        for data in array {
                            if let TextDrawAdjusted::Text(text) = data {
                                add_string(&text.data, &mut out, font);
                            }
                        }
                    }
                }
                Op::TextNewline => {
                    out.push('\n');
                }
                _ => {}
            }
        }
    }
    println!("{}", out);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_read_pdf() {
        read_text().unwrap();
    }
}
