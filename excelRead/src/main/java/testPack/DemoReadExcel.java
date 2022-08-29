package testPack;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class DemoReadExcel {

	public static void main(String[] args)throws IOException {
		FileInputStream file=new FileInputStream(new File("D:\\Details.xlsx"));
		XSSFWorkbook wb=new XSSFWorkbook(file);
		XSSFSheet sh=wb.getSheetAt(0);
		for(Row rw:sh) {
			System.out.println("  ");
			for(Cell cl:rw) {
				System.out.print(cl);
			}
			System.out.println("  ");
		}


	}

}
