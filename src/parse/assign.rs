use crate::parse::regex::match_project_no;
use crate::types::SummaryModelDocx;
use std::collections::HashMap;

pub fn parse_docx_table(content: Vec<String>) -> SummaryModelDocx {
    let mut summary = SummaryModelDocx::default();
    for (index, item) in content.iter().enumerate() {
        // 标题
        if item.contains("概要") && item.contains("Test Summary")
        {
            summary.title = item.clone().split("项目编号").next().unwrap().to_string();
        }
        // 项目编号
        if item.contains("项目编号")
        {
            summary.project_no = match_project_no(&item);
        }
        // 测试报告签发日期
        if item.contains("测试标准")
        {
            summary.base.test_date = content[index - 1].clone();
        }
        let field_mappings = HashMap::from([
            ("委托单位", &mut summary.base.consignor),
            ("生产单位", &mut summary.base.manufacturer),
            ("测试单位", &mut summary.base.test_lab),
            ("名称", &mut summary.base.cn_name),
            ("电芯类别", &mut summary.base.classification),
            ("型号", &mut summary.base.model),
            ("商标", &mut summary.base.trademark),
            ("额定电压", &mut summary.base.voltage),
            ("额定容量", &mut summary.base.capacity),
            ("额定能量", &mut summary.base.watt),
            ("外观", &mut summary.base.shape),
            ("质量", &mut summary.base.mass),
            ("锂含量", &mut summary.base.li_content),
            ("测试报告编号", &mut summary.base.test_report_no),
            ("测试标准", &mut summary.base.test_manual),
            ("高度模拟", &mut summary.base.test1),
            ("温度试验", &mut summary.base.test2),
            ("振动", &mut summary.base.test3),
            ("冲击", &mut summary.base.test4),
            ("外部短路", &mut summary.base.test5),
            ("撞击/挤压", &mut summary.base.test6),
            ("过度充电", &mut summary.base.test7),
            ("强制放电", &mut summary.base.test8),
            ("UN38.3.3.1(f)", &mut summary.base.un38_f),
            ("UN38.3.3.1(g)", &mut summary.base.un38_g),
            ("备注", &mut summary.base.note),
        ]);

        for (key, field) in field_mappings {
            if item.contains(key) {
                *field = content[index + 1].clone();
            }
        }
    }
    // 签发日期
    summary.issue_date = content[content.len() - 1].clone();
    summary
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{parse::read::parse_docx_text, reader::read_docx_content};

    #[test]
    fn test_parse_docx() {
        let text = read_docx_content(
            // r"C:\Users\29115\RustroverProjects\docx-rs\tests\test.docx",
            r"C:\Users\29115\Downloads\PEKGZ202412167637 概要.docx",
            vec!["word/document.xml".to_string()],
        );
        let content = parse_docx_text(&text.unwrap()[0].clone());
        println!("{}", content.clone().join("\n"));
        let summary = parse_docx_table(content);
        std::fs::write("test2.json", serde_json::to_string(&summary).unwrap()).unwrap();
    }
}