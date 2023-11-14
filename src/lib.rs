use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use anyhow::{anyhow,Result};
use zip::ZipArchive;

#[derive(Debug)]
pub struct ReplaceDocx {
    pub zip_file: ZipArchive<File>,
    pub content: String,
    pub links:String,
    pub headers:HashMap<String, String>,
    pub footers:HashMap<String, String>,
    pub images:HashMap<String, String>,
}


pub fn read_docx_file(file_path: &str) -> Result<ReplaceDocx> {
    if !file_path.ends_with("docx") {
        return Err(anyhow!("not docx file"));
    }
    let file = File::open(file_path)?;
    let mut zip_file = ZipArchive::new(file)?;
    let content = read_text(&mut zip_file)?;
    let (headers, footers) = read_header_footer(&mut zip_file)?;
    let links = read_link(&mut zip_file)?;
    let images = retrieve_images(&mut zip_file)?;


    Ok(ReplaceDocx{zip_file, content, headers, footers,links, images})
}

fn retrieve_images(file: &mut ZipArchive<File>) -> Result<HashMap<String, String>> {
    let mut images = HashMap::new();
    for i in 0..file.len(){
        let archive_file = file.by_index(i)?;
        let file_name = archive_file.name().to_string();
        if file_name.starts_with("word/media/") {
            images.insert(file_name.to_string(), "".to_string());
        }
    }
    Ok(images)
}

fn read_link(file: &mut ZipArchive<File>) -> Result<String> {
    let mut link = String::new();
    for i in 0..file.len() {
        let mut archive_file = file.by_index(i)?;
        if archive_file.name() == "word/_rels/document.xml.rels" {
            archive_file.read_to_string(&mut link)?;
        }
    }
    Ok(link.to_string())
}

fn read_header_footer(file: &mut ZipArchive<File>) -> Result<(HashMap<String,String>, HashMap<String, String>)> {
    let mut headers = HashMap::new();
    let footers = HashMap::new();
    for i in 0..file.len() {
        let mut archive_file = file.by_index(i)?;
        let file_name = archive_file.name().to_string();
        if file_name.contains("header") {
            let mut head = String::new();
            archive_file.read_to_string(&mut head)?;
            headers.insert(file_name.clone(), head);
        }
        if file_name.contains("footer") {
            let mut foot= String::new();
            archive_file.read_to_string(&mut foot)?;
            headers.insert(file_name, foot);
        }
    }
    Ok((headers, footers))
}

fn read_text(file: &mut ZipArchive<File>) -> Result<String> {
    let mut text = String::new();
    for i in 0..file.len(){
        let mut archive_file = file.by_index(i)?;
        let file_name = archive_file.name().to_string();
        if file_name == "word/document.xml"{
           archive_file.read_to_string(&mut text)?;
        }
    }
    Ok(text)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let result = read_docx_file("./hello.docx").unwrap();
        println!("{:?}", result);
        assert_eq!(4, 4);
    }
}
