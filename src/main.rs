mod bin{
    pub mod read_excel_as_csv_in_memory;
}


fn main() {
    let excel_file_path = "static/excel/Check Coordinates.xlsx";
    if let Ok(rows) = bin::read_excel_as_csv_in_memory::read_excel(excel_file_path) {
        if let Ok(csv_data) = bin::read_excel_as_csv_in_memory::write_csv_to_memory(&rows) {
            // Use the `csv_data` Vec<u8> as needed
            println!("CSV data:\n{}", String::from_utf8_lossy(&csv_data));
        } else {
            eprintln!("Error writing CSV to memory");
        }
    } else {
        eprintln!("Error reading Excel file");
    }
}