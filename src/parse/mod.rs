mod read;
mod assign;
mod regex;

pub use read::parse_docx_text;
pub use assign::parse_docx_table;
pub use regex::match_project_no;
