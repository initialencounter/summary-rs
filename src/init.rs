use lazy_static::lazy_static;
use regex::Regex;
use std::env;
use std::fs::File;
use std::io::{Read, Write};
use std::path::PathBuf;
use zip::write::{ExtendedFileOptions, FileOptions};
use zip::{ZipArchive, ZipWriter};

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

lazy_static! {
    static ref RE_IMAGE_EXTENT: Regex = Regex::new(r#"cx="(\d+)" cy="(\d+)""#).unwrap();
    static ref RE_IMAGE_BEHIND_DOCUMENT: Regex = Regex::new(r#"behindDoc="(\d)""#).unwrap();
}
pub fn get_signature_path() -> Result<String> {
    let exe_path = env::current_exe()?;
    let parent_path = exe_path.parent().ok_or("无法获取父目录")?;
    let signature_path_buf = parent_path.join("signature.png");
    let signature_path = signature_path_buf.to_str().ok_or("路径转换失败")?;
    Ok(signature_path.to_string())
}
pub fn read_file_to_buffer(file_path: &str) -> Result<Vec<u8>> {
    let mut file_content = Vec::new();
    File::open(PathBuf::from(file_path))?.read_to_end(&mut file_content)?;
    Ok(file_content)
}

fn change_title(content: String) -> String {
    let title = "锂电池/钠离子电池UN38.3试验概要";
    let content = content.replace("锂电池UN38.3试验概要", title);
    let content = content.replace("Lithium Battery Test Summary", "Test Summary");
    content.to_string()
}

fn change_test_info(content: String) -> String {
    let mut content = content.replacen("UN38.3.3(f)", "UN38.3.3.1(f)或/or\nUN38.3.3.2(d)", 1);
    content = content.replace("UN38.3.3(g)", "UN38.3.3.1(g) 或/or UN38.3.3.2(e)");
    content = content.replacen("UN38.3.3.1(f)", "dXNlIHN0ZDo6ZnM6OkZpbGU7CnVzZS1", 1);
    content = content.replacen("UN38.3.3.2(d)", "dXNlIHN0ZDo6ZnM6OkZpbGU7CnVzZS2", 1);
    content = content.replacen("UN38.3.3.1(f)", "UN38.3.3.1(g) ", 1);
    content = content.replacen("UN38.3.3.2(d)", "UN38.3.3.2(e)", 1);
    content = content.replace("dXNlIHN0ZDo6ZnM6OkZpbGU7CnVzZS1", "UN38.3.3.1(f)");
    content = content.replace("dXNlIHN0ZDo6ZnM6OkZpbGU7CnVzZS2", "UN38.3.3.2(d)");
    content.to_string()
}

pub fn set_image_size(content: String) -> Result<String> {
    let x = (3.5 * 360000.0) as i32;
    let y = (1.5 * 360000.0) as i32;
    let wp_extent = format!("cx=\"{}\" cy=\"{}\"", x, y);
    let content = RE_IMAGE_EXTENT.replace_all(&content, &wp_extent);
    Ok(content.to_string())
}

pub fn set_image_behind_document(content: String) -> Result<String> {
    let behind_doc = "behindDoc=\"1\"";
    let content = RE_IMAGE_BEHIND_DOCUMENT.replace_all(&content, &behind_doc.to_string());
    Ok(content.to_string())
}

pub fn modify_docx(input_path: &str) -> Result<()> {
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
            content = change_test_info(content);
            content = set_image_size(content)?;
            content = set_image_behind_document(content)?;
            zip_writer.start_file::<String, ExtendedFileOptions>(name, FileOptions::default())?;
            zip_writer.write_all(content.as_bytes())?;
        } else if name == "word/media/image1.png" {
            let signature_path = get_signature_path()?;
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
    output_file.write_all(&output_buffer)?;

    Ok(())
}
