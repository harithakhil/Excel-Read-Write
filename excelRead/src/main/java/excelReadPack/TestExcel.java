package excelReadPack;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class TestExcel {
	
	public static void main(String[] args) throws IOException {
		//to read file path
		FileInputStream file = new FileInputStream(new File("D:\\Student.xlsx"));
		//to create an object for workbook
		XSSFWorkbook wb = new XSSFWorkbook(file);
		//to create workbook sheet
		XSSFSheet sh = wb.getSheetAt(0);
		//for iterating row
		for(Row rw:sh) {
			//for iterating cells
			for(Cell cl:rw) {
				//for printing each cell
				System.out.print(cl);
			}
			System.out.println(" ");
		}
	}

}
