use calamine::{open_workbook, DataType, Reader, Xlsx};
use std::error::Error;
use std::io::{self, Cursor, Write};


pub fn read_excel(excel_file_path: &str) -> Result<Vec<Vec<String>>, Box<dyn Error>> {
    let mut workbook: Xlsx<_> = open_workbook(excel_file_path)?;
    let mut rows: Vec<Vec<String>> = Vec::new();

    if let Some(Ok(sheet)) = workbook.worksheet_range_at(0) {
        for row in sheet.rows() {
            let mut csv_row = Vec::new();
            for cell in row {
                if let DataType::String(value) = cell {
                    csv_row.push(value.to_owned());
                }
            }
            rows.push(csv_row);
        }
    }

    Ok(rows)
}

pub fn write_csv_to_memory(data: &[Vec<String>]) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut csv_writer = csv::Writer::from_writer(Vec::new());

    let max_fields = data.iter().map(|row| row.len()).max().unwrap_or(0);

    for row in data {
        let mut csv_row = row.clone();
        while csv_row.len() < max_fields {
            csv_row.push("".to_string()); // Add empty fields if the row has fewer fields
        }
        if csv_row.len() > max_fields {
            csv_row.truncate(max_fields); // Remove extra fields if the row has more fields
        }
        csv_writer.write_record(csv_row)?;
    }

    csv_writer.flush()?;
    let csv_data = csv_writer.into_inner()?;
    Ok(csv_data)
}