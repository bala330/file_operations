package org.in;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcelFile_CreateSheet {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		XSSFWorkbook workbook=new XSSFWorkbook();
		
		String Excelfilepath="E:\\SeleniumEx1\\FileOperations\\src\\test\\java\\org\\in\\Employe.Xlsx";

		FileOutputStream file=new FileOutputStream(Excelfilepath);
		
		XSSFSheet sheet=workbook.createSheet("Sheet1");
		
		System.out.println("sheets has been created");
		
		int numberofsheets=workbook.getNumberOfSheets();	
		
		System.out.println("Total number of sheets : " + numberofsheets);
		
		workbook.write(file);


	}

}
