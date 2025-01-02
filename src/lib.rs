mod init;
mod parse;
mod types;
mod reader;
mod modify_docx;

pub use types::*;
pub use init::modify_docx;
pub use reader::read_docx_content;
pub use parse::{parse_docx_table, parse_docx_text, match_project_no};
pub use modify_docx::modify_docx;