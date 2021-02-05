package org.tcs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo3 {
public static void main(String[] args) throws IOException {
	File file=new File("C:\\Users\\ASUS\\eclipse-workspace\\Demo\\Excel file\\Workbook.xlsx");
	FileInputStream Stream=new FileInputStream(file);
			Workbook workbook=new XSSFWorkbook(Stream);
       Sheet sheet = workbook.getSheet("Sheet1");
       Row row = sheet.getRow(1);
       Cell cell = row.getCell(1);
       String string = cell.getStringCellValue();
       if (string.equalsIgnoreCase("anusheya")) {
		cell.setCellValue("akila");
	}
       
       FileOutputStream stream1=new FileOutputStream(file);
       workbook.write(stream1);
       System.out.println("done");


}
}

