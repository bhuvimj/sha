package com.google.Practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XL {
	

	public static void pe() throws IOException {
		
	
		File f =new File("C:\\Users\\Bhuvi\\Desktop\\SN.xlsx");
		FileInputStream excel=new FileInputStream(f);
		Workbook wb= new XSSFWorkbook(excel);
		Sheet cS=wb.createSheet("stdents2");
		Row cR=cS.createRow(1);
		Cell createCell=cR.createCell(2);
		createCell.setCellValue("username");
		wb.getSheet("students2").getRow(1).createCell(2).setCellValue("username");
		FileOutputStream wr=new FileOutputStream(f);
		wb.write(wr);
		wb.close();
		
}
	
	public static void main(String[] args) throws IOException {

	}
		
}