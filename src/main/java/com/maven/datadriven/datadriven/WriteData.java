package com.maven.datadriven.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData {
	
	public static void excelDataWrite() throws InvalidFormatException, IOException {
		File f=new File("C:\\Manimegalai_course\\Workspace\\datadriven\\excelsheets\\excelsheet.xlsx");
		FileInputStream fis=new FileInputStream(f);
		Workbook wb=new XSSFWorkbook(fis);
		wb.createSheet("details").createRow(0).createCell(0).setCellValue("roll no");
		wb.getSheet("details").getRow(0).createCell(1).setCellValue("name");
		wb.getSheet("details").getRow(0).createCell(2).setCellValue("mobno");
		wb.getSheet("details").createRow(1).createCell(0).setCellValue("121");
		wb.getSheet("details").getRow(1).createCell(1).setCellValue("mani");
		wb.getSheet("details").getRow(1).createCell(2).setCellValue("9876543210");
		wb.getSheet("details").createRow(2).createCell(0).setCellValue("122");
		wb.getSheet("details").getRow(2).createCell(1).setCellValue("manju");
		wb.getSheet("details").getRow(2).createCell(2).setCellValue("9999998888");
		FileOutputStream fos=new FileOutputStream(f);
		wb.write(fos);
		System.out.println("successfully created");
		wb.close();
		
	}

	public static void main(String[] args) throws InvalidFormatException, IOException {
		excelDataWrite();

	}

}
