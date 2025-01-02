use std::fs::File;
use std::io::{Read, Write};
use zip::write::{ExtendedFileOptions, FileOptions};
use zip::{ZipArchive, ZipWriter};
mod utils;
use utils::*;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

pub fn modify_docx(input_path: &str, inspector: &str, width: f32, height: f32) -> Result<()> {
    // 先将整个文件读入内存
    let mut file_content = Vec::new();
    File::open(input_path)?.read_to_end(&mut file_content)?;

    // 从内存中读取zip文件
    let mut archive = ZipArchive::new(std::io::Cursor::new(&file_content))?;

    // 创建新的内存缓冲区用于存储修改后的文件
    let mut output_buffer = Vec::new();
    let mut zip_writer = ZipWriter::new(std::io::Cursor::new(&mut output_buffer));

    // 复制所有文件，修改document.xml
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_string();

        if name == "word/document.xml" {
            let mut content = String::new();
            file.read_to_string(&mut content)?;
            content = change_title(content);
            content = change_test_info(content, inspector);
            content = set_image_size(content, width, height)?;
            content = set_image_behind_document(content)?;
            content = set_page_margins(content)?;
            // content = set_specific_table_cell_width(content)?;
            zip_writer.start_file::<String, ExtendedFileOptions>(name, FileOptions::default())?;
            zip_writer.write_all(content.as_bytes())?;
        } else if name == "word/media/image1.png" {
            let signature_path = match get_signature_path() {
                Ok(path) => path,
                Err(_) => {
                    println!("无法获取签名路径");
                    continue;
                }
            };
            println!("signature_path: {}", signature_path);
            let buffer = read_file_to_buffer(&signature_path)?;
            zip_writer.start_file::<String, ExtendedFileOptions>(name, FileOptions::default())?;
            zip_writer.write_all(&buffer)?;
        } else {
            let mut buffer = Vec::new();
            file.read_to_end(&mut buffer)?;
            zip_writer.start_file::<String, ExtendedFileOptions>(name, FileOptions::default())?;
            zip_writer.write_all(&buffer)?;
        }
    }

    zip_writer.finish()?;

    // 直接写入原文件
    let mut output_file = File::create(input_path)?;
    match output_file.write_all(&output_buffer) {
        Ok(_) => {
            println!("修改成功");
        }
        Err(e) => {
            println!("无法写入文件: {}", e);
        }
    };

    Ok(())
}


#[cfg(test)]
mod tests {
    use super::*;
    use regex::Regex;

    #[test]
    fn test_modify_docx() {
        let _ = modify_docx("0.docx", "test", 5.58, 1.73);
    }

    #[test]
    fn test_replace_table_cell_width_with_const_string() {
        let content = r#"签名Signatory</w:t></w:r></w:p><w:p w14:paraId="4FC9CFFB"><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="center"/><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">职务 </w:t></w:r></w:p><w:p w14:paraId="0C3CBCD5"><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="center"/><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:szCs w:val="21"/></w:rPr><w:t>Title</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3272" w:type="dxa"/><w:gridSpan w:val="3"/><w:tcBorders><w:top w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:left w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:bottom w:val="single" w:color="auto" w:sz="12" w:space="0"/><w:right w:val="single" w:color="auto" w:sz="4" w:space="0"/></w:tcBorders><w:vAlign w:val="center"/></w:tcPr><w:p w14:paraId="333CBBDC"><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="center"/><w:rPr><w:rFonts w:eastAsia="华文行楷" w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251661312" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>330835</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>22225</wp:posOffset></wp:positionV><wp:extent cx="2008800" cy="622800"/><wp:effectExtent l="0" t="0" r="18415" b="10795"/><wp:wrapNone/><wp:docPr id="2" name="图片 2"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="2" name="图片 2"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2008800" cy="622800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p><w:p w14:paraId="4EAAA74E"><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="center"/><w:rPr><w:rFonts w:eastAsia="华文行楷" w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr></w:p><w:p w14:paraId="257FBBF8"><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="center"/><w:rPr><w:rFonts w:eastAsia="华文行楷" w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr></w:p><w:p w14:paraId="66BC216D"><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="center"/><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:szCs w:val="21"/></w:rPr><w:t>检验员Inspector：test</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1335" w:type="dxa"/><w:gridSpan w:val="2"/><w:tcBorders><w:top w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:left w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:bottom w:val="single" w:color="auto" w:sz="12" w:space="0"/><w:right w:val="single" w:color="auto" w:sz="4" w:space="0"/></w:tcBorders><w:vAlign w:val="center"/></w:tcPr><w:p w14:paraId="7750C236"><w:pPr><w:widowControl/><w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="center"/><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:szCs w:val="21"/></w:rPr><w:t>签发日期</w:t></w:r></w:p><w:p w14:paraId="47FC4B05"><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="center"/><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:szCs w:val="21"/></w:rPr><w:t>Issued Date</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3686" w:type="dxa"/><w:gridSpan w:val="5"/><w:tcBorders><w:top w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:left w:val="single" w:color="auto" w:sz="4" w:space="0"/><w:bottom w:val="single" w:color="auto" w:sz="12" w:space="0"/><w:right w:val="single" w:color="auto" w:sz="12" w:space="0"/></w:tcBorders><w:vAlign w:val="center"/></w:tcPr><w:sdt><w:sdtPr><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:szCs w:val="21"/></w:rPr><w:id w:val="147467335"/><w:placeholder><w:docPart w:val="{63add1fe-0aab-4e4c-ba56-35791ee4bace}"/></w:placeholder><w:date w:fullDate="2024-10-30T00:00:00Z"><w:dateFormat w:val="yyyy-MM-dd"/><w:lid w:val="en-US"/><w:storeMappedDataAs w:val="datetime"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr><w:sdtEndPr><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:szCs w:val="21"/></w:rPr></w:sdtEndPr><w:sdtContent><w:p w14:paraId="53F3CEBE"><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="center"/><w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorHAnsi"/><w:color w:val="auto"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr><w:t>2024-10-31</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl><w:p w14:paraId="6B74A4C4"><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:rPr><w:color w:val="000000" w:themeColor="text1"/><w:szCs w:val="21"/></w:rPr></w:pPr></w:p><w:sectPr><w:headerReference r:id="rId5" w:type="first"/><w:headerReference r:id="rId3" w:type="default"/><w:headerReference r:id="rId4" w:type="even"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1230" w:bottom="567" w:left="1230" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425" w:num="1"/><w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/></w:sectPr></w:body></w:document>"#;
        let re = Regex::new(r#"w:tcW w:w="1335"\sw:type="dxa""#).unwrap();
        let new_content = re.replace(content, "B0599DDA0D5649F4ACB5432F7C4003461").to_string();
        println!("{}", new_content);
    }

    #[test]
    fn test_set_specific_table_cell_width() {
      let text = "我有2个苹果和3个橘子。";
      let re = Regex::new(r"\d+").unwrap(); // 匹配一个或多个数字
  
      // 使用replace_all方法替换所有匹配的数字
      let result = re.replace(text, "X"); // 将数字替换为"X"
  
      println!("{}", result); // 输出: 我有X个苹果和X个橘子。
    }
}
