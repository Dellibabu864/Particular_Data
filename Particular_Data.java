package com.ParticularData2;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Particular_Data {
public static void main(String[] args) throws IOException {
	File f=new File("C:\\Users\\jeevi\\eclipse-workspace\\Maven_Project\\Project_Name\\Project.xlsx");
	FileInputStream fis= new FileInputStream(f);
	Workbook wb= new XSSFWorkbook(fis);
	Sheet s = wb.getSheet("Sheet1");
	Row r = s.getRow(2);
	Cell c = r.getCell(0); //string,numeric
	CellType ct = c.getCellType();
	if (ct.equals(CellType.STRING)) {
		
		String st = c.getStringCellValue();
		System.out.println(st);
		}
	else if (ct.equals(CellType.NUMERIC)) {
		
		double d = c.getNumericCellValue();
		int i=(int) d;//narrowing
		String value = String.valueOf(i);
		System.out.println(value);
		}
	wb.close();
}
}
