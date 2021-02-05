package org.tcs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo4 {
public static void main(String[] args) throws IOException {
	File file=new File("C:\\Users\\ASUS\\eclipse-workspace\\Demo\\Excel file\\Workbook.xlsx");
	FileInputStream stream=new FileInputStream(file);
	Workbook workbook=new XSSFWorkbook(stream);
	 Sheet createSheet = workbook.createSheet("data");
	 Row createRow = createSheet.createRow(1);
	 Cell createCell = createRow.createCell(0);
	 createCell.setCellValue("done");
	 
	 FileOutputStream Stream1=new FileOutputStream(file);
	 workbook.write(Stream1);
	 
	
	
}
}
