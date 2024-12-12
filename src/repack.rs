use quick_xml::events::{BytesEnd, BytesText};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use zip::write::ExtendedFileOptions;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

pub fn modify_docx(
    docx_data: &[u8],
    xml_path: &str,
    modify_fn: impl Fn(&[u8]) -> Vec<u8>,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    // 创建一个内存中的游标来读取docx数据
    let reader = Cursor::new(docx_data);
    let mut archive = ZipArchive::new(reader)?;

    // 创建一个新的内存buffer来存储修改后的docx
    let writer = Cursor::new(Vec::new());
    let mut zip_writer = ZipWriter::new(writer);

    // 存储所有文件的内容
    let mut files: HashMap<String, Vec<u8>> = HashMap::new();

    // 首先读取所有文件
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_string();
        let mut contents = Vec::new();
        file.read_to_end(&mut contents)?;
        files.insert(name, contents);
    }

    // 修改指定的XML文件
    if let Some(content) = files.get_mut(xml_path) {
        *content = modify_fn(content);
    }

    // 重新写入所有文件
    for (name, contents) in files {
        let options: FileOptions<'_, ExtendedFileOptions> =
            zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        zip_writer.start_file(&name, options)?;
        zip_writer.write_all(&contents)?;
    }

    // 完成写入并返回结果
    let writer = zip_writer.finish()?;
    Ok(writer.into_inner())
}

// 使用示例
pub fn example_usage() -> Result<(), Box<dyn std::error::Error>> {
    // 读取docx文件
    let docx_data = std::fs::read("tests/test 概要.docx")?;

    // 定义修改XML的函数
    let modify_xml = |xml_content: &[u8]| {
        // 这里是修改XML的逻辑
        // 这只是一个示例，实际使用时需要根据具体需求修改
        let mut reader = Reader::from_reader(xml_content);
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.name().as_ref() == b"w:t" => {
                    writer.write_event(Event::Start(e.clone())).unwrap();

                    // 读取原文本内容
                    let mut original_text = String::new();
                    loop {
                        match reader.read_event_into(&mut buf) {
                            Ok(Event::Text(e)) => {
                                original_text = e.unescape().unwrap().into_owned();
                                break;
                            }
                            Ok(Event::End(ref e)) if e.name().as_ref() == b"w:t" => {
                                writer.write_event(Event::End(e.clone())).unwrap();
                                break;
                            }
                            Ok(event) => writer.write_event(event).unwrap(),
                            Err(e) => {
                                panic!("Error at position {}: {:?}", reader.buffer_position(), e)
                            }
                        }
                    }

                    // 根据条件判断是否需要替换
                    let new_text = if original_text.contains("2024-10-30") {
                        "2024-10-31"
                    } else {
                        &original_text
                    };

                    // 写入文本（可能是新文本或原文本）
                    writer
                        .write_event(Event::Text(BytesText::new(new_text)))
                        .unwrap();

                    // 写入结束标签
                    if !original_text.is_empty() {
                        writer.write_event(Event::End(BytesEnd::new("w:t"))).unwrap();
                    }
                }
                Ok(Event::End(ref e)) if e.name().as_ref() == b"w:t" => {
                    // 写入结束标签
                    writer.write_event(Event::End(e.clone())).unwrap();
                }
                Ok(Event::Eof) => break,
                Ok(event) => {
                    writer.write_event(event).unwrap();
                }
                Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            }
            buf.clear();
        }

        writer.into_inner().into_inner()
    };

    // 修改docx
    let modified_docx = modify_docx(&docx_data, "word/document.xml", modify_xml)?;

    // 保存修改后的文件
    std::fs::write("output.docx", modified_docx)?;

    Ok(())
}
