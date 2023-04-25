package com.maven.datadriven.datadriven;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven_1 {
	
	public static void getParticularData() throws InvalidFormatException, IOException {
		File f=new File("C:\\Manimegalai_course\\Workspace\\datadriven\\excelsheets\\excelsheet.xlsx");
		Workbook wb=new XSSFWorkbook(f);
		Sheet sheetAt=wb.getSheetAt(0);
		Row row=sheetAt.getRow(1);
		Cell cell=row.getCell(1);
		String stringCellValue=cell.getStringCellValue();
		System.out.println(stringCellValue);
	}

	public static void main(String[] args) throws InvalidFormatException, IOException {
		getParticularData();

	}

}
