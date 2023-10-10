package org.in;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		String Excelfilepath="E:\\SeleniumEx1\\FileOperations\\src\\test\\java\\org\\in\\Student.Xlsx";
		FileInputStream input=new FileInputStream(Excelfilepath);
		
		XSSFWorkbook workbook=new XSSFWorkbook(input);
		  XSSFSheet sheet=workbook.getSheet("Student Info");
		int rows = sheet.getLastRowNum();
		int columns=sheet.getRow(1).getLastCellNum();
		for(int r=0;r<=rows;r++) {
			XSSFRow row=sheet.getRow(r);
			for(int c=0;c<columns;c++) {
				 XSSFCell cell=row.getCell(c);
				 
				switch (cell.getCellType()) {
				case STRING:
                System.out.print(cell.getStringCellValue());
                break;
				case NUMERIC: 
					System.out.print(cell.getNumericCellValue());
					break;
			    case BOOLEAN:
			    	System.out.print(cell.getBooleanCellValue());
			    	break;
				}
				System.out.print(" | ");
				}
		}
		System.out.println();
	}

}
