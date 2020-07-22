package excel2;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;

public class Excel2 {
	
	public static void main(String[] args) {
		Workbook workbook = new HSSFWorkbook();
		
		//create a sheet with the name "Eggs"
		Sheet sheet = workbook.createSheet("Eggs");
		
		//create the row for our cell. It starts counting from 0, so this is row 1 in Excel
		Row row = sheet.createRow(1);
		//create the cell in column number 'A'. So now we have a cell in 'A1'
		Cell cell1 = row.createCell(3);
		
		//add some content to that cell
		cell1.setCellValue("Hi there");
		
		
		//a more compact version of creating cells in a specified location
		Cell cell2 = sheet.createRow(0).createCell(0); //Cell 'A2'
		cell2.setCellValue("Weslley Felix");
		
		try {
			FileOutputStream output = new FileOutputStream("Test1.xls");
			workbook.write(output);
			output.close();
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
}