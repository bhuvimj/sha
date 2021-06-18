package com.google.Practice;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Exl {
	private static void pd() throws IOException
	{
		File f= new File ("C:\\Users\\Bhuvi\\Desktop\\SN.xlsx");
	    FileInputStream excel = new FileInputStream(f);
		@SuppressWarnings("resource")
		Workbook wb=new XSSFWorkbook(excel); 
	    Sheet sheetAt= wb.getSheetAt(0);	
	    int rows =sheetAt.getPhysicalNumberOfRows();
	    System.out.println("no.of rows:"+rows);
	    for(int i=0;i<rows;i++){
	    Row r=sheetAt.getRow(i);
	    int cells=r.getPhysicalNumberOfCells();
	    for(int j=0;j<cells;j++){
	    Cell cell=r.getCell(j);
	    CellType celltype = cell.getCellType();
	   // System.out.println(cell.getCellType());
	    if(celltype.equals(CellType.STRING)) {
	    	String sv=cell.getStringCellValue();
	    	System.out.println(sv);
	    }
	    else if(celltype.equals(CellType.NUMERIC)) {
	    	double vs=cell.getNumericCellValue();
	    	int k=(int)vs;
	    	System.out.println(k);
	    }
	    }  
	}
	}
	public static void main(String[] args) throws IOException {
		pd();	
}
}