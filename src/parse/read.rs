use quick_xml::events::Event;
use quick_xml::Reader;

pub fn parse_docx_text(content: &str) -> Vec<String> {
    let mut reader = Reader::from_str(content);
    let mut buf = Vec::new();
    let mut path: Vec<(String, usize)> = Vec::new(); // 存储(标签名, 索引)
    let mut counts: std::collections::HashMap<String, usize> = std::collections::HashMap::new(); // 记录每个标签的计数
    let mut output = Vec::<String>::new();
    let mut last_path_str: String = "w:document[1]/w:body[1]/w:p[1]//".to_string();
    let mut last_text: String = "".to_string();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();
				// 处理 w:date 标签
                if e.name().as_ref() == b"w:date" {
                    // 获取 fullDate 属性
                    if let Some(attr) = e.attributes().find(|a| {
                        let a = a.clone().unwrap();
                        a.key.as_ref() == b"w:fullDate"
                    }) {
                        match attr {
							Ok(attr) => {
								let date =  String::from_utf8_lossy(attr.value.as_ref());
								if let Some(date_only) = date.split('T').next() {
									last_text = date_only.to_string();
									output.push(last_text.clone());
								}
							},
							Err(e) => {
								println!("{}", e);
							}

						}
                    }
                }
                // 更新该标签的计数
                let count = counts.entry(tag_name.clone()).or_insert(0);
                *count += 1;

                // 将标签名和索引加入路径栈
                path.push((tag_name, *count));

                if e.name().as_ref() == b"w:t" {
                    if let Ok(Event::Text(e)) = reader.read_event_into(&mut buf) {
                        // 打印带索引的完整路径
                        let path_str = path
                            .iter()
                            .map(|(tag, idx)| {
                                if tag == "w:t" || tag == "w:r" || tag == "w:p" {
                                    "".to_string()
                                } else {
                                    format!("{}[{}]", tag, idx)
                                }
                            })
                            .collect::<Vec<_>>()
                            .join("/");
                        let text = e.unescape().unwrap_or_default().to_string();
                        if last_path_str.is_empty() {
                            last_path_str = path_str.clone();
                        }
                        if path_str.clone() != last_path_str.clone() {
                            output.push(last_text.clone());
                            last_text = text.clone();
                        } else {
                            last_text = format!("{}{}", last_text, text);
                        }
                        last_path_str = path_str;
                    }
                }
            }
            Ok(Event::End(_)) => {
                path.pop();
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                println!("解析错误: {}", e);
                break;
            }
            _ => (),
        }
        buf.clear();
    }
    output.push(last_text.clone());
    output
}