// Learn more about Tauri commands at https://tauri.app/develop/calling-rust/

// use base64::decode;
use base64::{decode, DecodeSliceError};
use calamine::{Cell, Data, DataType, Reader, Xlsx};
use chrono::{Duration, NaiveDate};
// use docx_rs::*;
// use docx::{Document, Docx, Paragraph, Run};
use base64;
use serde::de::IntoDeserializer;
use serde::{Deserialize, Serialize};
use serde_json;
use std::fs;
use std::fs::File;
use std::io::Cursor;
use std::io::Read;
use std::io::Write;

// use docx_rust::document::Paragraph;
// use docx_rust::DocxFile;
// use std::io::{self, Write};
#[derive(Debug, Serialize, Deserialize)]
struct DataItem {
    org_name: String,
    org_code: String,
    check_month: String,
    check_day: String,
    rec_month: String,
    rec_day: String,
    stp_month: String,
    stp_day: String,
}

#[tauri::command]
fn greet(name: &str) -> String {
    format!("Hello, {}! You've been greeted from Rust!", name)
}

#[tauri::command]
fn choose_word_tpl(name: &str) -> String {
    format!("Hello, {}! You've been greeted from Rust!", name)
}

#[tauri::command]
fn get_docx_bs64(docx_path: &str) -> String {
    // 打开 Word 文件
    let mut file = File::open(docx_path).expect("无法打开文件");

    // 读取文件的二进制数据
    let mut buffer = Vec::new();
    file.read_to_end(&mut buffer).expect("无法读取文件");

    // 将二进制数据转换为 Base64 编码
    let encoded = base64::encode(&buffer);
    // 返回 Base64 编码的字符串
    return encoded;
}

// 写文件
#[tauri::command]
fn save_docx_bs64(docx_path: &str, bs64: &str) -> Result<(), String> {
    // 将 Base64 编码的字符串解码为二进制数据
    let decoded = decode(bs64).map_err(|e| format!("无法解码 Base64 字符串: {}", e))?;

    // 将二进制数据写入文件
    let mut file = File::create(docx_path).expect("无法创建文件");
    println!("doc_path:{}", docx_path);
    file.write_all(&decoded).expect("无法写入文件");
    Ok(())
}

#[tauri::command]
fn process_file(docx_path: &str, save_path: &str, excel_file: &str) -> Result<String, String> {
    if !fs::metadata(&docx_path).is_ok() {
        return Err(format!("Word 文件不存在:{}", docx_path).into());
    }
    if !fs::metadata(&save_path).is_ok() {
        return Err(format!("保存路径不存在:{}", save_path).into());
    }

    if excel_file.is_empty() {
        return Err("Excel 文件不能为空".into());
    }
    // println!("excel_file:{}", excel_file);
    // match decode(&excel_file) {
    //     Ok(content) => {
    //         //base64解码
    //     }
    // }
    let datas = read_excel_from_base64(&excel_file);

    let mut ret_list: Vec<DataItem> = Vec::new();
    // 循环datas,并根据datas生成docx，保存到save_path
    for data in datas.unwrap() {
        // 生成docx
        // 保存到save_path
        println!("data: {:?}", data);
        //构建替换模板
        // let replacements = vec![
        //     ("{{org_name}}", data.0.as_str()),
        //     ("{{org_no}}", data.1.as_str()),
        //     ("{{check_month}}", data.2.format("%m").to_string().as_str()),
        //     ("{{check_day}}", data.2.format("%d").to_string().as_str()),
        //     ("{{rec_month}}", data.3.format("%d").to_string().as_str()),
        //     ("{{rec_day}}", data.3.format("%d").to_string().as_str()),
        //     ("{{stp_month}}", data.4.format("%d").to_string().as_str()),
        //     ("{{stp_day}}", data.4.format("%d").to_string().as_str()),
        // ];
        // 列表字典
        ret_list.push(DataItem {
            org_name: data.0,
            org_code: data.1,
            check_month: data.2.format("%m").to_string(),
            check_day: data.2.format("%d").to_string(),
            rec_month: data.3.format("%m").to_string(),
            rec_day: data.3.format("%d").to_string(),
            stp_month: data.4.format("%m").to_string(),
            stp_day: data.4.format("%d").to_string(),
        });
        //替换模板
        // generate_docx(docx_path, &replacements, save_path);
    }
    // 转成json字符传
    let json_str = serde_json::to_string(&ret_list).unwrap();
    println!("json_str:{}", json_str);
    return Ok(json_str);
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_shell::init())
        .invoke_handler(tauri::generate_handler![
            greet,
            choose_word_tpl,
            process_file,
            get_docx_bs64,
            save_docx_bs64
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

// 读取
fn read_excel_from_base64(
    base64_content: &str,
) -> Result<Vec<(String, String, NaiveDate, NaiveDate, NaiveDate)>, Box<dyn std::error::Error>> {
    let base64_content = base64_content.split(',').nth(1).unwrap_or(base64_content);
    // 解码 Base64 字符串为二进制数据
    let binary_data = decode(base64_content)?;

    // 使用 Cursor 将二进制数据转换为可读的流
    let mut cursor = Cursor::new(binary_data);

    // 打开 Excel 文件
    let mut workbook: Xlsx<_> = Xlsx::new(&mut cursor)?;

    let mut extracted_data = Vec::new();
    let base_date = NaiveDate::from_ymd(1899, 12, 30); // Excel's base date
                                                       // 获取第一个工作表
    if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
        // 遍历行
        for row in range.rows().skip(1) {
            // // 打印每一行的单元格
            // for cell in row {
            //     print!("{:?}\t", cell);
            // }

            let unit_name = row
                .get(0)
                .and_then(DataType::as_string)
                .unwrap_or("".into())
                .to_string();
            let number = row
                .get(1)
                .and_then(DataType::as_string)
                .unwrap_or("".into())
                .to_string();
            let current_date = row
                .get(2)
                .and_then(DataType::as_i64)
                .and_then(|cell| Some(base_date + chrono::Duration::days(cell)))
                .unwrap_or(NaiveDate::from_ymd(1900, 1, 1));
            let rectification_date = row
                .get(3)
                .and_then(DataType::as_i64)
                .and_then(|cell| Some(base_date + chrono::Duration::days(cell)))
                .unwrap_or(NaiveDate::from_ymd(1900, 1, 1));
            let stamp_date = row
                .get(4)
                .and_then(DataType::as_i64)
                .and_then(|cell| Some(base_date + chrono::Duration::days(cell)))
                .unwrap_or(NaiveDate::from_ymd(1900, 1, 1));

            // Add the extracted data to the vector
            extracted_data.push((
                unit_name,
                number,
                current_date,
                rectification_date,
                stamp_date,
            ));
        }
    } else {
        println!("First sheet not found or is empty");
    }

    Ok(extracted_data)
}

// fn excel_date_to_naive_date(Data: &dyn Data) -> Option<NaiveDate> {
//     let _ = Data;
//     if let DataType::DateTime(date) = cell.data_type() {
//         return date.and_hms(0, 0, 0).naive_utc().to_local().ok();
//     }
//     None
// }

// 替换word模板数据并生成新word
// fn replace_placeholders_in_docx(
//     template_path: &str,
//     output_path: &str,
//     replacements: &[(String, String)],
// ) -> Result<(), Box<dyn std::error::Error>> {
//     // Read the template file
//     let mut file = File::open(template_path)?;
//     let mut buffer = Vec::new();
//     file.read_to_end(&mut buffer)?;

//     // Load the document
//     let mut doc = Docx::from_bytes(&buffer)?;

//     // Iterate over paragraphs and replace placeholders
//     for paragraph in doc.paragraphs_mut() {
//         for run in paragraph.runs_mut() {
//             if let Some(text) = run.text_mut() {
//                 for (placeholder, replacement) in replacements {
//                     *text = text.replace(placeholder, replacement);
//                 }
//             }
//         }
//     }

//     // Save the modified document
//     doc.write_to_file(output_path)?;

//     println!("Document saved at {}", output_path);
//     Ok(())
// }

// fn generate_docxfile(
//     docx_path: &str,
//     relacements: &[(String, String)],
//     output: &str,
// ) -> Result<(), Box<dyn std::error::Error>> {
//     let docx = DocxFile::from_file(docx_path).unwrap();
//     let mut docx = docx.parse().unwrap();

//     println!("Document parsed successfully")
//     // 获取所有段落
//     // let mut paragraphs_mut = docx;

//     // let filename = relacements.0.clone().1.clone();
//     // let outputpath = format!("{}\\{}.docx", output, filename);
//     // for paragraph in doc.paragraphs_mut() {
//     //     let text = paragraph.text();
//     //     let mut new_text = text.clone();

//     //     // 替换标记
//     //     for (placeholder, value) in &replacements {
//     //         new_text = new_text.replace(placeholder, value);
//     //     }

//     //     // 清空段落并添加新的文本
//     //     paragraph.clear();
//     //     paragraph.add_run(Run::new().text(new_text));
//     // }
//     // doc.save(outputpath).unwrap();
//     Ok(())
// }
