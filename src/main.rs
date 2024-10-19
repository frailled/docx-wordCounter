use std::fs::File;
use std::io::{self, BufReader};
use zip::read::ZipArchive;
use xml::reader::{EventReader, XmlEvent};

fn main() {
    // Prompt the user for a .docx file path
    println!("Enter the .docx file path:");

    // Get the input from the user
    let mut file_path = String::new();
    io::stdin().read_line(&mut file_path).expect("Failed to read input");

    // Trim the newline character from the input
    let file_path = file_path.trim();

    // Try to open the .docx file as a zip archive
    let file = File::open(file_path).expect("Could not open the file");
    let reader = BufReader::new(file);
    let mut zip = ZipArchive::new(reader).expect("Could not read .docx file as zip archive");

    // Find the word/document.xml file inside the .docx archive
    let mut document_xml = zip.by_name("word/document.xml").expect("word/document.xml not found in .docx");

    // Read the document.xml file and parse the text
    let parser = EventReader::new(&mut document_xml);
    let mut text_content = String::new();
    let mut line_count = 0;
    let mut in_paragraph = false;

    for event in parser {
        match event {
            Ok(XmlEvent::StartElement { name, .. }) => {
                if name.local_name == "p" {
                    // Start of a new paragraph
                    line_count += 1;  // Count paragraph as a new line
                    in_paragraph = true; // Indicate we're inside a paragraph
                }
                if name.local_name == "r" {
                    // Start of a run
                    in_paragraph = true; // Still inside a paragraph
                }
                if name.local_name == "br" {
                    // Count line breaks within paragraphs
                    if in_paragraph {
                        line_count += 1; // Each <w:br> counts as a new line
                    }
                }
            }
            Ok(XmlEvent::EndElement { name }) => {
                if name.local_name == "p" {
                    // End of a paragraph
                    in_paragraph = false; // Reset paragraph state
                }
            }
            Ok(XmlEvent::Characters(text)) if in_paragraph => {
                // Collect text for word and character counting only if inside a paragraph
                text_content.push_str(&text);
                text_content.push(' ');
            }
            _ => {}
        }
    }

    // Now we have the full text, we can count words and characters
    let word_count = text_content.split_whitespace().count();
    let char_count = text_content.chars().count();

    // Output the results
    println!("Lines: {}", line_count);
    println!("Words: {}", word_count);
    println!("Characters: {}", char_count);
}
