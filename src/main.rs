use serde_json::Value;
use std::io;
use std::{
    error::Error,
    fs::{self, File},
    io::BufReader,
};
use xlsxwriter::{FormatAlignment, FormatColor, FormatUnderline, Workbook};

//表示一门语言
#[derive(Debug)]
struct Lan {
    name: String,
    texts: Vec<String>,
}

fn get_languages(path: &str) -> Result<Vec<Lan>, Box<dyn Error>> {
    let mut languages = Vec::new();
    let files = fs::read_dir(path);
    let files = match files {
        Ok(files) => files,
        Err(e) => return Err(Box::new(e)),
    };

    for entry in files {
        let path = entry?.path();
        let name = path.file_stem().unwrap().to_str().unwrap().to_string();
        if path.extension().unwrap().ne("json") {
            continue;
        }

        let file = File::open(path)?;
        let reader = BufReader::new(file);
        let v: Value = serde_json::from_reader(reader)?;

        let mut texts = vec![];
        for ele in v.as_object().unwrap().iter() {
            if name == "en" {
                texts.push(ele.0.to_string());
            } else {
                texts.push(ele.1.as_str().unwrap().to_string());
            }
        }
        languages.push(Lan { name, texts });
    }

    Ok(languages)
}

fn export_excel(languages: Vec<Lan>, path: &str) -> Result<(), Box<dyn Error>> {
    let workbook = Workbook::new(path);
    let title_format = workbook
        .add_format()
        .set_font_color(FormatColor::Blue)
        .set_underline(FormatUnderline::Single);
    let first_col_format = workbook
        .add_format()
        .set_font_color(FormatColor::Green)
        .set_align(FormatAlignment::Left)
        .set_text_wrap();

    let col_format = workbook
        .add_format()
        .set_align(FormatAlignment::Left)
        .set_text_wrap();

    let mut sheet: xlsxwriter::Worksheet = workbook.add_worksheet(None)?;
    for (col, lan) in languages.iter().enumerate() {
        sheet.write_string(0, col as u16, &lan.name, Some(&title_format))?;
        for (row, text) in lan.texts.iter().enumerate() {
            if !&text.trim_end().is_empty() {
                if col == 0 {
                    sheet.write_string(
                        (row + 1) as u32,
                        col as u16,
                        &text,
                        Some(&first_col_format),
                    )?;
                } else {
                    sheet.write_string((row + 1) as u32, col as u16, &text, Some(&col_format))?;
                }
            }
        }
    }
    sheet.set_column(0, 0, 80.0, None)?;

    workbook.close()?;
    Ok(())
}

fn get_input(tip: &str, default: &str) -> Result<String, Box<dyn Error>> {
    println!("{}", tip);
    let mut path = String::new();
    io::stdin().read_line(&mut path)?;
    path = path.trim().to_string();
    if path.is_empty() {
        path = String::from(default.trim());
    }
    Ok(path)
}

fn main() {
    //获取多语言
    let mut lang: Vec<Lan>;
    loop {
        let path = get_input("输入多语言目录（默认：.\\）: ", ".\\").expect("读取输入异常");
        lang = match get_languages(&path) {
            Ok(r) => {
                println!("获取多语言成功");
                r
            }
            Err(e) => {
                println!("获取json出错{}", e);
                continue;
            }
        };

        if lang.len() > 0 {
            break;
        }
    }
    //写入excel
    const DEFAULT_OUTPUT_NAME: &str = "lan.xlsx";
    let path = get_input(
        &format!("输入导出位置（默认:{}）: ", DEFAULT_OUTPUT_NAME),
        DEFAULT_OUTPUT_NAME,
    )
    .expect("读取异常");
    export_excel(lang, &path).expect("写入excel失败");

    println!("导出成功");
    let _ = get_input("按任意键结束", "");
}
