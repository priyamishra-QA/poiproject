package POI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeexcel {
	

	public static void main(String[] args) throws Exception {
		try {
		// TODO Auto-generated method stub
		FileInputStream fis = new FileInputStream ("C:\\Users\\Sanjay\\Desktop\\testdata1.xlsx");

		/*Workbook wb = WorkbookFactory.create(fis);

		Sheet sh = wb.getSheet("Sheet1");
		Row row = sh.getRow(1);
		Cell cel = row.createCell(3);
		cel.setCellType(CellType.STRING);

		FileOutputStream fos = new FileOutputStream("C:\\Users\\Sanjay\\Desktop\\testdata1.xlsx");

		cel.setCellValue("pass");
		wb.write(fos);
System.out.println("completed");*/
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sh1= wb.getSheetAt(0);
		System.out.println(sh1.getRow(0).getCell(0).getStringCellValue());
		System.out.println(sh1.getRow(0).getCell(1).getStringCellValue());
		System.out.println(sh1.getRow(1).getCell(0).getStringCellValue());
		System.out.println(sh1.getRow(1).getCell(1).getStringCellValue());
		System.out.println(sh1.getRow(2).getCell(0).getStringCellValue());
		System.out.println(sh1.getRow(2).getCell(1).getStringCellValue());
		//To write or input value to excel sheet
		FileOutputStream fos = new FileOutputStream("C:\\Users\\Sanjay\\Desktop\\testdata1.xlsx");
		
		sh1.getRow(0).createCell(2).setCellValue("priya");
		sh1.getRow(1).createCell(2).setCellValue("nimo");
		sh1.getRow(2).createCell(2).setCellValue("tunu");
		wb.write(fos);
		fos.close();
	}catch (Exception e) {
		System.out.println(e.getMessage());
		System.out.println("completed");
	}
	
	
		
		
	}
	

}
