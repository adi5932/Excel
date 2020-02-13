package com.main;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExData {

	
	public static void main(String args[]) throws IOException {
		
		ReadExData ReadData= new ReadExData();
		String output= ReadData.ReadCellData(1, 1).toString();
		System.out.print(output);
		
	}
	public String ReadCellData(int Row, int Col)
	{
		
		Workbook wb= null;
		try {
			FileInputStream fis= new FileInputStream("D:\\Nya folder\\ExcelFetch\\src\\ExcelExample.xlsx");
		wb= new XSSFWorkbook(fis);
		}
		catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		catch(IOException e) {
			e.printStackTrace();
		}
		org.apache.poi.ss.usermodel.Sheet sheet= wb.getSheetAt(0);
		Row r= sheet.getRow(Row);
		Cell cell= r.getCell(Col);
		
	String Res= cell.getStringCellValue();
		return Res;
	}
}




























