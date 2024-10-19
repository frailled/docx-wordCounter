use std::fs::File;
use std::io::{self, BufReader};
use zip::read::ZipArchive;
use xml::reader::{EventReader, XmlEvent};

fn main() {
    println!("Enter the .docx file path:");

    let mut file_path = String::new();
    io::stdin().read_line(&mut file_path).expect("Failed to read input");

    let file_path = file_path.trim();

    let file = File::open(file_path).expect("Could not open the file");
    let reader = BufReader::new(file);
    let mut zip = ZipArchive::new(reader).expect("Could not read .docx file as zip archive");

    let mut document_xml = zip.by_name("word/document.xml").expect("word/document.xml not found in .docx");

    let parser = EventReader::new(&mut document_xml);
    let mut text_content = String::new();
    let mut line_count = 0;
    let mut in_paragraph = false;

    for event in parser {
        match event {
            Ok(XmlEvent::StartElement { name, .. }) => {
                if name.local_name == "p" {
                    line_count += 1;
                    in_paragraph = true; 
                }
            }
            Ok(XmlEvent::EndElement { name }) => {
                if name.local_name == "p" {
                    in_paragraph = false; 
                }
            }
            Ok(XmlEvent::Characters(text)) if in_paragraph => {
                text_content.push_str(&text);
                text_content.push(' ');
            }
            _ => {}
        }
    }

    let word_count = text_content.split_whitespace().count();
    let char_count = text_content.chars().count();

    println!("Lines: {}", line_count);
    println!("Words: {}", word_count);
    println!("Characters: {}", char_count);
}
