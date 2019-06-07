String excelfile="SampBind6_copy.xlsm";

PRFile file = new PRFile(tools.findPage("pxProcess").getString("pxServiceExportPath")+excelfile);

try{
   PRInputStream inputStream = new PRInputStream(file);
   org.apache.poi.ss.usermodel.Workbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(inputStream);
   org.apache.poi.ss.usermodel.Sheet sheet=workbook.getSheet("Benefits Package 1");
   org.apache.poi.ss.usermodel.Cell cell2Update = sheet.getRow(8).getCell(7);
   org.apache.poi.ss.usermodel.Cell cell9Update = sheet.getRow(8).getCell(8);
   org.apache.poi.ss.usermodel.CellType  type =cell2Update.getCellType();
   oLog.infoForced(cell2Update.getStringCellValue()+" "+type.toString());
   cell2Update.setCellValue("Existing");
   cell9Update.setCellValue("HMO");
   // update the file
   PROutputStream out= new PROutputStream(tools.findPage("pxProcess").getString("pxServiceExportPath")+"Updated.xlsm");
   workbook.write(out);
   out.flush();
   workbook.close();
   out.close();
  
  /*java.util.Iterator<org.apache.poi.ss.usermodel.Row> rowIterator = sheet.iterator();
    while (rowIterator.hasNext()) {
      org.apache.poi.ss.usermodel.Row row = rowIterator.next();
      java.util.Iterator<org.apache.poi.ss.usermodel.Cell> cellIterator = row.cellIterator();
       while (cellIterator.hasNext()) {
         org.apache.poi.ss.usermodel.Cell cell = cellIterator.next();
         org.apache.poi.ss.usermodel.CellType  type =cell.getCellType();
         //oLog.infoForced(String.valueOf(cell.getColumnIndex() ));
          oLog.infoForced(cell.getStringCellValue()+" "+type.toString());
          //oLog.infoForced();
       }
    } */
   //org.apache.poi.ss.usermodel.Cell cell2Update = sheet.getRow(9).getCell(8);
   //oLog.infoForced(cell2Update.getStringCellValue() );
  
   //cell2Update.setCellValue("Existing");
  /*
  //
  PRInputStream inputStream = new PRInputStream(file);
  org.apache.poi.ss.usermodel.Workbook workbook=org.apache.poi.ss.usermodel.WorkbookFactory.create(inputStream);
  org.apache.poi.ss.usermodel.Sheet sheet=workbook.getSheetAt(0);
  org.apache.poi.ss.usermodel.Cell cell2Update = sheet.getRow(9).getCell(8);
  cell2Update.setCellValue("Existing");
  // write file
  //PROutputStream out= new PROutputStream(tools.findPage("pxProcess").getString("pxServiceExportPath")+"Updated.xlsm");
  //workbook.write(out);
 // workbook.close();
  //out.close();
  */
   
}catch(Exception e){
  e.printStackTrace();
}