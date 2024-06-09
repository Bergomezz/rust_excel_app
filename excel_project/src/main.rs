extern crate calamine;

use calamine::{open_workbook_auto, Reader};
use std::{error::Error, fs};

fn main() -> Result<(), Box<dyn Error>> {

  let path = "C:\\Users\\berna\\OneDrive\\Área de Trabalho\\teste1.xlsx";
  let mut workbook = open_workbook_auto(path)?;

  let mut items_uniq_wb: Vec<_> = Vec::new();

  if let Ok(range) = workbook.worksheet_range("Plan1") {
  for row in range.rows() {
      items_uniq_wb.push(row);
      println!("{:?}", items_uniq_wb);
  }
  }




  let mut vector_entries: Vec<_> = Vec::new();
  let entries = fs::read_dir("C:\\Users\\berna\\OneDrive\\Documentos")?;
  let mut items_from_vector: Vec<_> = Vec::new();


  for entry in entries {
      let entry = entry?;
      let path = entry.path();
      if path.is_file() && path.extension().and_then(|s| s.to_str()) == Some("xlsx") {
        vector_entries.push(path);
      }
  }


  for path in &vector_entries {
      let path_str = path.to_str().unwrap();
      let mut workbooks = open_workbook_auto(path_str).unwrap();

      if let Ok(range) = workbooks.worksheet_range("Plan1"){
        for rows in range.rows(){
          println!("{:?}", rows);
        }
      }
    };





  Ok(())
}
