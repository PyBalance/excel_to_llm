use calamine::{ open_workbook, Reader, Xlsx };
use std::error::Error;
use eframe::egui::{ self, FontData, FontDefinitions, FontFamily };
use std::sync::mpsc;
use std::thread;
use strum_macros::{EnumIter, Display};

fn main() -> Result<(), eframe::Error> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default().with_inner_size([800.0, 600.0]),
        ..Default::default()
    };
    eframe::run_native(
        "Excel Analyzer",
        options,
        Box::new(|cc| Ok(Box::new(ExcelAnalyzerApp::new(cc))))
    )
}

struct ExcelAnalyzerApp {
    file_paths: Vec<String>,
    output: String,
    tx: mpsc::Sender<String>,
    rx: mpsc::Receiver<String>,
    rows_to_display: String,
    output_format: OutputFormat,
    is_analyzing: bool,
}

#[derive(Debug, Clone, PartialEq, EnumIter, Display)]
enum OutputFormat {
    Markdown,
    XML,
    PlainText,
}

impl ExcelAnalyzerApp {
    fn new(cc: &eframe::CreationContext<'_>) -> Self {
        Self::configure_fonts(&cc.egui_ctx);
        let (tx, rx) = mpsc::channel();
        Self {
            file_paths: Vec::new(),
            output: String::new(),
            tx,
            rx,
            rows_to_display: "5".to_string(), // Default to 5 rows
            output_format: OutputFormat::Markdown,
            is_analyzing: false,
        }
    }

    fn configure_fonts(ctx: &egui::Context) {
        let mut fonts = FontDefinitions::default();
        
        // Load your custom font
        fonts.font_data.insert(
            "NotoSansCJK".to_owned(),
            FontData::from_static(include_bytes!("../assets/NotoSansSC-VariableFont_wght.ttf")),
        );

        // Set as the first font for proportional text
        fonts.families.get_mut(&FontFamily::Proportional).unwrap()
            .insert(0, "NotoSansCJK".to_owned());

        // Set as the first font for monospace text
        fonts.families.get_mut(&FontFamily::Monospace).unwrap()
            .insert(0, "NotoSansCJK".to_owned());

        ctx.set_fonts(fonts);
    }
}

impl Default for ExcelAnalyzerApp {
    fn default() -> Self {
        let (tx, rx) = mpsc::channel();
        Self {
            file_paths: Vec::new(),
            output: String::new(),
            tx,
            rx,
            rows_to_display: "5".to_string(), // Default to 5 rows
            output_format: OutputFormat::Markdown,
            is_analyzing: false,
        }
    }
}

impl eframe::App for ExcelAnalyzerApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("Excel Analyzer");

            ui.horizontal(|ui| {
                if ui.button("Add Excel File").clicked() {
                    if let Some(path) = rfd::FileDialog::new()
                        .add_filter("Excel Files", &["xlsx", "xls"])
                        .pick_file()
                    {
                        let path_str = path.to_str().unwrap().to_string();
                        self.file_paths.push(path_str);
                    }
                }

                ui.add(egui::TextEdit::singleline(&mut self.rows_to_display)
                    .desired_width(50.0)
                    .hint_text("Rows"));

                if let Ok(rows) = self.rows_to_display.parse::<usize>() {
                    self.rows_to_display = rows.to_string();
                }

                egui::ComboBox::from_label("Output Format")
                    .selected_text(format!("{:?}", self.output_format))
                    .show_ui(ui, |ui| {
                        ui.selectable_value(&mut self.output_format, OutputFormat::Markdown, "Markdown");
                        ui.selectable_value(&mut self.output_format, OutputFormat::XML, "XML");
                        ui.selectable_value(&mut self.output_format, OutputFormat::PlainText, "Plain Text");
                    });
            });

            ui.group(|ui| {
                ui.label("Selected Files:");
                let mut remove_index = None;
                for (index, path) in self.file_paths.iter().enumerate() {
                    ui.horizontal(|ui| {
                        ui.label(format!("{}: {}", index + 1, path));
                        if ui.button("Remove").clicked() {
                            remove_index = Some(index);
                        }
                    });
                }
                if let Some(index) = remove_index {
                    self.file_paths.remove(index);
                }
            });

            if ui.button("Analyze Files").clicked() && !self.is_analyzing {
                self.output.clear();
                self.is_analyzing = true;
                
                let tx = self.tx.clone();
                let file_paths = self.file_paths.clone();
                let rows_to_display = self.rows_to_display.clone();
                let output_format = self.output_format.clone();
                thread::spawn(move || {
                    for file_path in file_paths {
                        match process_excel_file(&file_path, &rows_to_display, &output_format) {
                            Ok(output) => {
                                tx.send(output).unwrap();
                            }
                            Err(e) => {
                                tx.send(format!("Error processing {}: {}", file_path, e)).unwrap();
                            }
                        }
                    }
                    tx.send("Analysis complete".to_string()).unwrap();
                });
            }

            if self.is_analyzing {
                ui.add(egui::Spinner::new());
            }

            if let Ok(new_output) = self.rx.try_recv() {
                if new_output == "Analysis complete" {
                    self.is_analyzing = false;
                } else {
                    self.output.push_str(&new_output);
                    self.output.push_str("\n\n");
                }
            }

            egui::ScrollArea::vertical().show(ui, |ui| {
                ui.add(egui::TextEdit::multiline(&mut self.output).desired_width(f32::INFINITY));
            });
        });
    }
}

fn process_excel_file(file_path: &str, rows_to_display: &str, output_format: &OutputFormat) -> Result<String, Box<dyn Error>> {
    let mut workbook: Xlsx<_> = open_workbook(file_path)?;
    let mut output = String::new();

    let rows_to_display = rows_to_display.parse::<usize>().unwrap_or(5);

    for (sheet_index, sheet_name) in workbook.sheet_names().iter().enumerate() {
        let sheet = match workbook.worksheet_range(sheet_name) {
            Ok(sheet) => sheet,
            Err(e) => {
                return Err(format!("Sheet not found: {} - {}", sheet_name, e).into());
            }
        };

        let headers: Vec<String> = sheet
            .rows()
            .next()
            .ok_or("Sheet is empty")?
            .iter()
            .map(|cell| cell.to_string())
            .collect();

        match output_format {
            OutputFormat::Markdown => {
                output.push_str(&format!("# Excel File Name: {}\n\n", file_path));
                output.push_str(&format!("## Sheet {}:\n\n", sheet_index + 1));
                output.push_str(&format!("### Sheet Name: {}\n\n", sheet_name));
                output.push_str("### Headers:\n");
                for (i, header) in headers.iter().enumerate() {
                    output.push_str(&format!("- {}: {}\n", i + 1, header));
                }
                output.push_str("\n### Sample Data:\n\n");

                // Create markdown table header
                output.push_str("|");
                for header in &headers {
                    output.push_str(&format!(" {} |", header));
                }
                output.push_str("\n|");
                for _ in &headers {
                    output.push_str(" --- |");
                }
                output.push_str("\n");

                // Add table rows
                for row in sheet.rows().skip(1).take(rows_to_display) {
                    output.push_str("|");
                    for cell in row {
                        output.push_str(&format!(" {} |", cell.to_string()));
                    }
                    output.push_str("\n");
                }
                output.push_str("\n");
            },
            OutputFormat::XML => {
                output.push_str(&format!("<excel-file name=\"{}\">\n", file_path));
                output.push_str(&format!("  <sheet index=\"{}\" name=\"{}\">\n", sheet_index + 1, sheet_name));
                output.push_str("    <headers>\n");
                for (i, header) in headers.iter().enumerate() {
                    output.push_str(&format!("      <header index=\"{}\">{}</header>\n", i + 1, header));
                }
                output.push_str("    </headers>\n");
                output.push_str("    <sample-data>\n");
                for row in sheet.rows().skip(1).take(rows_to_display) {
                    output.push_str("      <row>\n");
                    for cell in row {
                        output.push_str(&format!("        <cell>{}</cell>\n", cell.to_string()));
                    }
                    output.push_str("      </row>\n");
                }
                output.push_str("    </sample-data>\n");
                output.push_str("  </sheet>\n");
                output.push_str("</excel-file>\n\n");
            },
            OutputFormat::PlainText => {
                output.push_str(&format!("Excel File Name: {}\n\n", file_path));
                output.push_str(&format!("Sheet {}:\n", sheet_index + 1));
                output.push_str(&format!("Sheet Name: {}\n\n", sheet_name));
                output.push_str("Headers:\n");
                for (i, header) in headers.iter().enumerate() {
                    output.push_str(&format!("{}: {}\n", i + 1, header));
                }
                output.push_str("\nSample Data:\n\n");
                for row in sheet.rows().skip(1).take(rows_to_display) {
                    for cell in row {
                        output.push_str(&format!("{}\t", cell.to_string()));
                    }
                    output.push_str("\n");
                }
                output.push_str("\n");
            },
        }
    }

    Ok(output)
}
