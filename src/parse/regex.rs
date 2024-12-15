use lazy_static::lazy_static;
use regex::Regex;
lazy_static! {
    static ref RE_PROJECT_NO: Regex = Regex::new(r"([PSAR]EKGZ[0-9]{12})").unwrap();
}
pub fn match_project_no(content: &str) -> String {
    let matches: Vec<String> = RE_PROJECT_NO
        .captures_iter(&content)
        .filter_map(|cap| cap[1].parse::<String>().ok())
        .collect();
    if matches.is_empty() {
        return "".to_string();
    } else {  
        matches[0].to_string()
    }
}
