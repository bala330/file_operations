package org.in;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		

	

		XSSFWorkbook workbook= new XSSFWorkbook();
		XSSFSheet sheet= workbook.createSheet("Student Info");
		
		Object studata[][]= {  {"StudentID","StudentName","StudentEmailID"},
				               {108,"Bala","bala458@gmail.com"},
				{109,"Hari","hariram234@gmail.com"},
				{110,"Arun","arunkumar1985@gmail.com"},
				{111,"Smith","smith345@gmail.com"},
				{112,"Scott","scott36@gmail.com"}
					};
		
		int rows=studata.length;
		int columns=studata[0].length;
		System.out.println(rows);
		System.out.println(columns);
		
		
		for(int r=0;r<rows;r++) {
			
			XSSFRow row=sheet.createRow(r);
		for(int c=0;c<columns;c++) {
			 XSSFCell cell =row.createCell(c);
			 Object value= studata[r][c];
			 
			 if (value instanceof String) 
				 cell.setCellValue((String)value);
			 if (value instanceof Integer) 
				 cell.setCellValue((Integer)value);
			 if (value instanceof Boolean) 
				 cell.setCellValue((Boolean)value);
		}
		}
		String filepath="E:\\SeleniumEx1\\FileOperations\\src\\test\\java\\org\\in\\Student.Xlsx";
		FileOutputStream outputstream=new FileOutputStream(filepath);
		workbook.write(outputstream);
		outputstream.close();
		
		System.out.println("Student.xls file is written successfully");
        				
		

	}

}

