package org.in;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteaData {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		

	

		XSSFWorkbook workbook= new XSSFWorkbook();
		XSSFSheet sheet= workbook.createSheet("Sheet1");
	

		
		Object Empdata[][]= {  {"NAME","AGE","EmailID"},
				               {"John Doe",30,"john@test.com"},
				               {"Jane Doe",28,"john@test.com"},
				               {"Bob Smith",35,"jacky@example.com"},
				               {"Swapnil",37,"joe@example.com"}
			                   
					};
		
		int rows=Empdata.length;
		int columns=Empdata[0].length;
		System.out.println(rows);
		System.out.println(columns);
		
		
		for(int r=0;r<rows;r++) {
			
			XSSFRow row=sheet.createRow(r);
		for(int c=0;c<columns;c++) {
			 XSSFCell cell =row.createCell(c);
			 Object value= Empdata[r][c];
			 
			 if (value instanceof String) 
				 cell.setCellValue((String)value);
			 if (value instanceof Integer) 
				 cell.setCellValue((Integer)value);
			 if (value instanceof Boolean) 
				 cell.setCellValue((Boolean)value);
		}
		}
		String filepath="E:\\SeleniumEx1\\FileOperations\\src\\test\\java\\org\\in\\Employe.Xlsx";
		FileOutputStream outputstream=new FileOutputStream(filepath);
		workbook.write(outputstream);
		outputstream.close();


	}

}

