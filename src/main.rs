use native_dialog::{FileDialog, MessageDialog, MessageType};
use std::{collections::HashMap, error::Error, path::Path};

fn main() -> Result<(), Box<dyn Error>> {
    let template_xlsx = FileDialog::new()
        .add_filter("Spreadsheet", &["xlsx"])
        .show_open_single_file()?;

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

    let CEFR = row0.get("CEFR").expect(":");
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

    for pdf in glob::glob("*results*.pdf")
        .expect("Unable to find pdf files with filename including results")
    {
        let doc = Document::load(pdf.unwrap()).unwrap();
        let mut font_maps = HashMap::new();
        for page_id in doc.page_iter() {
            let content = doc.get_and_decode_page_content(page_id).unwrap();
            let fonts = doc.get_page_fonts(page_id);
            for (name, font) in fonts.into_iter() {
                font_maps.entry(name).or_insert_with_key(|name| {
                    let to_unicode = font
                        .get_deref(b"ToUnicode", &doc)
                        .unwrap()
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
            let mut student_details: Details = Default::default();
            let mut grades: Vec<String> = Vec::with_capacity(16);
            let mut subject_label: Vec<String> = Vec::with_capacity(16);
            for operation in content.operations.iter() {
                match operation.operator.as_str() {
                    "ET" => {
                        let gleaned_text = str.trim();
                        // println!("{pos:?} {gleaned_text} {state:?} {anchor:?} ");

                        match state {
                            State::CandidateId if anchor == pos.y => {
                                student_details.candidate_id = gleaned_text.to_string();
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
                                c.as_str()
                                    .unwrap()
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

            let row = id_col
                .get(student_details.candidate_id.as_str())
                .expect("Unable to fetch id");

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
        }
    }
    for pdf in
        glob::glob("*cefr*.pdf").expect("Unable to find pdf files with filename including results")
    {
        let doc = Document::load(pdf.unwrap()).unwrap();
        let mut font_maps = HashMap::new();
        for page_id in doc.page_iter() {
            let content = doc.get_and_decode_page_content(page_id).unwrap();
            let fonts = doc.get_page_fonts(page_id);
            for (name, font) in fonts.into_iter() {
                font_maps.entry(name).or_insert_with_key(|name| {
                    let to_unicode = font
                        .get_deref(b"ToUnicode", &doc)
                        .unwrap()
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
            let mut student_details: Details_CEFR_Quirky = Default::default();
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
                                if student_details.cefr.is_none() {
                                    student_details.cefr = Some(gleaned_text.to_string());
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
                                    )
                                    .unwrap()
                                    .matches(x.as_str()) =>
                                    {
                                        student_details.candidate_id = gleaned_text.to_string();
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
                                c.as_str()
                                    .unwrap()
                                    .chunks_exact(2)
                                    .into_iter()
                                    .map(|c| {
                                        current_font
                                            .get(&(u16::from_be_bytes([c[0], c[1]]) as u32))
                                            .unwrap()
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

            let row = id_col
                .get(student_details.candidate_id.as_str())
                .expect(format!("Unable to fetch id {}", student_details.candidate_id).as_str());

            let cell_value = sheet.get_cell_value_mut((*CEFR, *row));
            cell_value.set_value(student_details.cefr.unwrap());
        }
    }
    umya_spreadsheet::writer::xlsx::write(&book, &Path::new("example_output.xlsx"))
        .expect("Unable to write to xlsx file");

    Ok(())
}
