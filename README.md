## Data experimental tools using Rust programming language

- __Read excel as csv in memory__
The module `src/bin/read_excel_as_csv_in_memory.rs` reads and excel file, without including vba or formulas. Then it writes the data as csv in a memory object. This is convenient for using the csv object with [Polars](https://www.pola.rs), as it does not have a native Excel reader.