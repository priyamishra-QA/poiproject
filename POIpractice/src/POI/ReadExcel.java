\package POI;

import java.io.FileInputStream;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		FileInputStream fis = new FileInputStream ("C:\\Users\\Sanjay\\Desktop\\testdata1.xlsx");
		//Workbook wb= WorkbookFactory.create(fis);
		XSSFWorkbook  wb = new XSSFWorkbook(fis);
		Sheet sh = wb.getSheet("Sheet1");
		Row row = sh.getRow(3);
		String Username = row.getCell(0).getStringCellValue();
		String Password = row.getCell(1).getStringCellValue();
		System.out.println(Username);
		System.out.println(Password);


	}
	 //This method is to write in the Excel cell, Row num and Col num are the parameters
	 
    public static void setCellData(String Result,  int RowNum, int ColNum) throws Exception {

      try{

          Row  = ExcelWSheet.getRow(RowNum);

Cell = Row.getCell(ColNum, Row.RETURN_BLANK_AS_NULL);

if (Cell == null) {

Cell = Row.createCell(ColNum);

Cell.setCellValue(Result);

} else {

Cell.setCellValue(Result);

}

         // Constant variables Test Data path and Test Data file name

          FileOutputStream fileOut = new FileOutputStream(Constant.Path_TestData + Constant.File_TestData);

          ExcelWBook.write(fileOut);

          fileOut.flush();

fileOut.close();

}catch(Exception e){

throw (e);

///another approach.......................................................................
/*@Test
48
 public void ReadData() throws IOException
49
 {
50
     // Import excel sheet.
51
     File src=new File("C:\\Users\\Admin\\Desktop\\TestData.xls");
52
      
53
     // Load the file.
54
     FileInputStream finput = new FileInputStream(src);
55
      
56
     // Load he workbook.
57
    workbook = new HSSFWorkbook(finput);
58
      
59
     // Load the sheet in which data is stored.
60
     sheet= workbook.getSheetAt(0);
61
      
62
     for(int i=1; i&lt;=sheet.getLastRowNum(); i++)
63
     {
64
         // Import data for Email.
65
         cell = sheet.getRow(i).getCell(1);
66
         cell.setCellType(Cell.CELL_TYPE_STRING);
67
         driver.findElement(By.id("login-email")).sendKeys(cell.getStringCellValue());
68
          
69
         // Import data for password.
70
         cell = sheet.getRow(i).getCell(2);
71
         cell.setCellType(Cell.CELL_TYPE_STRING);
72
         driver.findElement(By.id("login-password")).sendKeys(cell.getStringCellValue());
73
           // // Write data in the excel..........................................................
75
       FileOutputStream foutput=new FileOutputStream(src);
76
         
77
        // Specify the message needs to be written.
78
        String message = "Data Imported Successfully.";
79
         
80
        // Create cell where data needs to be written.
81
        sheet.getRow(i).createCell(3).setCellValue(message);
82
          
83
        // Specify the file in which data needs to be written.
84
        FileOutputStream fileOutput = new FileOutputStream(src);
85
         
86
        // finally write content
87
        workbook.write(fileOutput);
88
         
89
         // close the file
90
        fileOutput.close();
      
74
        }
75
  }*/
;