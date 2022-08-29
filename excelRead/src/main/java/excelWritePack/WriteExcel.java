package excelWritePack;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws IOException {
		//create a blank workbook
		XSSFWorkbook wb=new XSSFWorkbook();
		//create a sheet for workbook
		XSSFSheet sh=wb.createSheet("student details");
		//create a row object
		XSSFRow rw;
		Map<String, Object[]> data=new HashMap<String, Object[]>();
		data.put("1", new Object[] {"SlNo","Name","Mark"});
		data.put("2", new Object[] {"101","Haritha","50"});
		data.put("3", new Object[] {"102","Archana","60"});
		data.put("4", new Object[] {"103","Joseph","70"});
		Set<String> keyid=data.keySet();
		int rowid=0;
		for(String key:keyid) {
			rw=sh.createRow(rowid++);
			Object[] objarr=data.get(key);
			int cellid=0;
			for(Object obj:objarr) {
				Cell cell=rw.createCell(cellid++);
				cell.setCellValue((String)obj);
			}
		}
		 FileOutputStream out = new FileOutputStream(new File("D:\\StDetails1.xlsx"));
		 wb.write(out);
		 out.close();
		 System.out.println("Excel file is write successfully");
	}

}
