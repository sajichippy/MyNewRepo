package pack;

import java.io.FileInputStream;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	
	XSSFSheet sh;  //sheet1
	
	public Excel() throws IOException  {
		FileInputStream f = new FileInputStream("C:\\Users\\HP\\OneDrive\\Desktop\\ExcelRead.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(f);
		sh = wb.getSheet("Sheet1");
	}
	
	public String readData(int i, int j) {  //0 0
		Row r = sh.getRow(i);  //0
		Cell c = r.getCell(j);  //0
		int celltype = c.getCellType();   //0 or 1
		switch(celltype) {
		case Cell.CELL_TYPE_NUMERIC:
		{
			double d = c.getNumericCellValue();  //120
			return String.valueOf(d);
			
		}
		
		case  Cell.CELL_TYPE_STRING:
		{
			return c.getStringCellValue();
		}
		
		
		
		
	}
		return null;}
}
