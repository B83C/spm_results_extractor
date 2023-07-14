use encoding::Encoding;
use encoding::all::UTF_16BE;
use encoding::DecoderTrap;
use lopdf::Document;
use native_dialog::{FileDialog, MessageDialog, MessageType};
use std::{collections::HashMap, error::Error, path::Path};

use std::path::PathBuf;

struct Pos<T> {
    x: T,
    y: T,
}

impl<T> Pos<T> {
    fn new(x: T, y: T) -> Self {
        Self { x, y }
    }
}

enum State {
    CandidateId,
    Grade,
    SubjectCode,
    CEFR,
    Unknown,
}

fn execute() -> Result<(), Box<dyn Error>> {

    MessageDialog::new()
        .set_title("Notice")
        .set_text("Please select the template spreadsheet file for data insertion")
        .show_alert()
        .unwrap();

    let template_xlsx = FileDialog::new()
        .add_filter("Spreadsheet", &["xlsx"])
        .show_open_single_file().expect("Unable to get the template spreadsheet file for inserting data");

    let mut book = umya_spreadsheet::reader::xlsx::lazy_read(template_xlsx.unwrap().as_ref())?;
    let sheet = book
        .get_sheet_by_name_mut("All")
        .expect("Unable to get sheet 'All'");
    let row0: HashMap<_, _> = sheet
        .get_collection_by_row(&1)
        .into_iter()
        .map(|x| {
            (
                x.get_raw_value().to_string(),
                *x.get_coordinate().get_col_num(),
            )
        })
        .collect();

    let CEFR = row0.get("CEFR").expect("Unable to get the CEFR column");
    let id_col: HashMap<_, _> = sheet
        .get_collection_by_column(
            &sheet
                .get_collection_by_row(&2)
                .iter()
                .find(|x| {
                    x.get_raw_value()
                        .to_string()
                        .to_lowercase()
                        .contains("giliran")
                })
                .expect("Unable to find Angka Giliran column")
                .get_coordinate()
                .get_col_num(),
        )
        .into_iter()
        .map(|x| {
            (
                x.get_raw_value().to_string(),
                *x.get_coordinate().get_row_num(),
            )
        })
        .collect();

    MessageDialog::new()
        .set_title("Notice")
        .set_text("Please select the pdf result files, you may choose not to")
        .show_alert()
        .unwrap();

    if let Ok(results_files) = FileDialog::new()
        .add_filter("PDF Doc", &["pdf"])
        .show_open_multiple_file(){
        for pdf in results_files
        {
            let doc = Document::load(pdf)?;
            let mut font_maps = HashMap::new();
            for page_id in doc.page_iter() {
                let content = doc.get_and_decode_page_content(page_id)?;
                let fonts = doc.get_page_fonts(page_id);
                for (name, font) in fonts.into_iter() {
                    font_maps.entry(name).or_insert_with_key(|name| {
                        let to_unicode = font
                            .get_deref(b"ToUnicode", &doc).unwrap()
                            .as_stream()
                            .expect(format!("Unable to dereference font for {:?}", name).as_str())
                            .decompressed_content()
                            .expect("Unable to decompress stream for font");
                        let mapping = adobe_cmap_parser::get_unicode_map(to_unicode.as_ref())
                            .expect("Unable to get unicode cmap");
                        mapping
                    });
                }
                let mut current_font = None;
                let mut str = String::new();
                let mut pos = Pos::new(0., 0.);
                let mut anchor = 0.;
                let mut state = State::Unknown;
                let mut candidate_id = None;
                let mut grades: Vec<String> = Vec::with_capacity(16);
                let mut subject_label: Vec<String> = Vec::with_capacity(16);
                for operation in content.operations.iter() {
                    match operation.operator.as_str() {
                        "ET" => {
                            let gleaned_text = str.trim();

                            match state {
                                State::CandidateId if anchor == pos.y => {
                                    candidate_id = Some(gleaned_text.to_string());
                                }
                                State::SubjectCode if anchor == pos.x => {
                                    if gleaned_text.parse::<u32>().is_ok() {
                                        subject_label.push(gleaned_text.to_string());
                                    }
                                }
                                State::Grade if anchor == pos.x => {
                                    grades.push(gleaned_text.to_string())
                                }
                                _ => {
                                    (state, anchor) = match gleaned_text.to_lowercase() {
                                        x if x.contains("angka") && x.contains("giliran") => {
                                            (State::CandidateId, pos.y)
                                        }
                                        x if x.contains("gred") => (State::Grade, pos.x),
                                        x if x.contains("kod") => (State::SubjectCode, pos.x),
                                        _ => (state, anchor),
                                    };
                                }
                            }
                            str.clear();
                        }
                        "Tf" => {
                            let font_name = operation.operands[0]
                                .as_name()
                                .expect("Unable to get name of Tf");
                            current_font = Some(
                                font_maps
                                    .get(font_name)
                                    .expect("Unable to fetch preloaded fonts"),
                            );
                        }
                        "Tj" => {
                            if let Some(current_font) = current_font {
                                for c in operation.operands.iter() {
                                    c.as_str()?
                                        .chunks_exact(2)
                                        .into_iter()
                                        .map(|c| {
                                            current_font
                                                .get(&(u16::from_be_bytes([c[0], c[1]]) as u32))
                                                .expect("Text not aligned evenly, please contact the developer")
                                        })
                                        .for_each(|x| {
                                            str.push_str(
                                                UTF_16BE
                                                    .decode(x, DecoderTrap::Strict)
                                                    .unwrap_or("".to_owned())
                                                    .as_str(),
                                            );
                                        });
                                }
                            }
                        }
                        "Tm" => {
                            pos = Pos::new(
                                operation.operands[4].as_f32().unwrap_or(
                                    operation.operands[4].as_i64().unwrap_or_default() as f32,
                                ),
                                operation.operands[5].as_f32().unwrap_or(
                                    operation.operands[5].as_i64().unwrap_or_default() as f32,
                                ),
                            );
                        }
                        _ => {}
                    }
                }

                if let Some(ref candidate_id) = candidate_id {
                    let row = id_col.get(candidate_id).expect("Unable to fetch id");

                    subject_label
                        .into_iter()
                        .filter_map(|x| {
                            if let Some(value) = row0.get(x.as_str()) {
                                Some(value)
                            } else {
                                eprintln!(
                                    "Subject {} not found, please add it into the column",
                                    x.as_str()
                                );
                                None
                            }
                        })
                        .zip(grades.into_iter())
                        .for_each(|(subject_col, grade)| {
                            let cell_value = sheet.get_cell_value_mut((*subject_col, *row));
                            cell_value.set_value(grade);
                        });
                } else {
                    println!(
                        "Unable to get candidate ID from page {} of {}",
                        "test", "test"
                    );
                    todo!();
                };
            }
        }    
    }

    MessageDialog::new()
        .set_title("Notice")
        .set_text("Please select the cefr pdf result files, you may choose not to ")
        .show_alert()
        .unwrap();

    if let Ok(cefr_results_files )= FileDialog::new()
        .add_filter("PDF Doc", &["pdf"])
        .show_open_multiple_file() {
        for pdf in cefr_results_files
        {
            let doc = Document::load(pdf)?;
            let mut font_maps = HashMap::new();
            for page_id in doc.page_iter() {
                let content = doc.get_and_decode_page_content(page_id)?;
                let fonts = doc.get_page_fonts(page_id);
                for (name, font) in fonts.into_iter() {
                    font_maps.entry(name).or_insert_with_key(|name| {
                        let to_unicode = font
                            .get_deref(b"ToUnicode", &doc).unwrap()
                            .as_stream()
                            .expect(format!("Unable to dereference font for {:?}", name).as_str())
                            .decompressed_content()
                            .expect("Unable to decompress stream for font");
                        let mapping = adobe_cmap_parser::get_unicode_map(to_unicode.as_ref())
                            .expect("Unable to get unicode cmap");
                        mapping
                    });
                }
                let mut current_font = None;
                let mut str = String::new();
                let mut pos = Pos::new(0, 0);
                let mut anchor = 0;
                let mut state = State::Unknown;
                let mut candidate_id = None;
                let mut cefr_grade = None;
                for operation in content.operations.iter() {
                    match operation.operator.as_str() {
                        "ET" => {
                            let gleaned_text = str.trim();
                            // println!("{pos:?} {gleaned_text} {state:?} {anchor:?} ");

                            match state {
                                // State::Name if anchor == pos.y => {
                                //     student_details.name = gleaned_text.to_string();
                                // }
                                // State::IcNo if anchor == pos.y => {
                                //     student_details.ic_no = gleaned_text.to_string();
                                // }
                                State::CEFR if anchor == pos.y => {
                                    if cefr_grade.is_none() {
                                        cefr_grade = Some(gleaned_text.to_string());
                                    }
                                }
                                // State::NumberOfGrades if anchor == pos.y => {}
                                _ => {
                                    (state, anchor) = match gleaned_text.to_lowercase() {
                                        x if x.contains("tahap")
                                            && x.contains("cefr")
                                            && x.contains(":") =>
                                        {
                                            (State::CEFR, pos.y)
                                        }
                                        x if glob::Pattern::new(
                                            "wl[0-9][0-9][0-9]d[0-9][0-9][0-9]",
                                        )?
                                        .matches(x.as_str()) =>
                                        {
                                            candidate_id = Some(gleaned_text.to_string());
                                            (state, anchor)
                                        }
                                        _ => (state, anchor),
                                    };
                                }
                            }
                            str.clear();
                        }
                        "Tf" => {
                            let font_name = operation.operands[0]
                                .as_name()
                                .expect("Unable to get name of Tf");
                            current_font = Some(
                                font_maps
                                    .get(font_name)
                                    .expect("Unable to fetch preloaded fonts"),
                            );
                        }
                        "Tj" => {
                            if let Some(current_font) = current_font {
                                for c in operation.operands.iter() {
                                    c.as_str()?
                                        .chunks_exact(2)
                                        .into_iter()
                                        .map(|c| {
                                            current_font
                                                .get(&(u16::from_be_bytes([c[0], c[1]]) as u32)).unwrap()
                                        })
                                        .for_each(|x| {
                                            str.push_str(
                                                UTF_16BE
                                                    .decode(x, DecoderTrap::Strict)
                                                    .unwrap_or("".to_owned())
                                                    .as_str(),
                                            );
                                        });
                                }
                            }
                        }
                        "Tm" => {
                            pos = Pos::new(
                                operation.operands[4].as_i64().unwrap_or(
                                    operation.operands[4].as_f32().unwrap_or_default() as i64,
                                ),
                                operation.operands[5].as_i64().unwrap_or_default(),
                            );
                        }
                        _ => {}
                    }
                }

                if let Some(ref candidate_id) = candidate_id {
                    let row = id_col
                        .get(candidate_id)
                        .expect(format!("Unable to fetch id {}", candidate_id).as_str());

                    let cell_value = sheet.get_cell_value_mut((*CEFR, *row));
                    cell_value.set_value(cefr_grade.expect(format!("Unable to get cefr_grade for candidate : {}", candidate_id).as_str()));
                }
            }
        }    
    }

    let output_file = FileDialog::new()
        .add_filter("Spreadsheet", &["xlsx"])
        .show_save_single_file().expect("Please tell me the path to save the output file").unwrap_or(PathBuf::from("output.xlsx"));

    umya_spreadsheet::writer::xlsx::write(&book, output_file)
        .expect("Unable to write to xlsx file");

    Ok(())
}

fn main() {
    if let Err(e) = execute() {
        MessageDialog::new()
            .set_title("Error occurred, please contact the developer")
            .set_text(&e.to_string())
            .show_alert()
            .unwrap();
    }
}
