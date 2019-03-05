package com.demo.JDBCExcel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcelSheet {
	private static String[] columns = {"Roll No"};
	public static void main(String[] args) throws IOException {
		//xssf workbook
				Workbook workbook1 = new XSSFWorkbook();
				Sheet sheet1 = workbook1.createSheet("Random1");
						
				Font headerFont1 = workbook1.createFont();
				headerFont1.setBold(true);
				headerFont1.setFontHeightInPoints((short)16);
				headerFont1.setColor(IndexedColors.GOLD.getIndex());
				
				CellStyle headerCellStyle1 = workbook1.createCellStyle();
				headerCellStyle1.setFont(headerFont1);
				
				Row headerRow1 = sheet1.createRow(0);
				
				for(int i=0; i < columns.length;i++)
				{
					Cell cell = headerRow1.createCell(i);
					cell.setCellValue(columns[i]);
					cell.setCellStyle(headerCellStyle1);
				}
				
				int rowNum1 = 1;
				while(rowNum1 < 200000)
				{
					Row row = sheet1.createRow(rowNum1++);
					
					row.createCell(0).setCellValue(rowNum1-1);
					
				}
			
				
				try {
				//FileOutputStream fileout1 = new FileOutputStream("XSSF-File.xlsx");
					FileOutputStream fileout1 = new FileOutputStream("XSSF-File.xls");
				workbook1.write(fileout1);
				fileout1.close();
				}
				catch(FileNotFoundException f)
				{
					f.getMessage();
				}
				finally
				{
				System.out.println("Data written successfully...!");
				System.out.println("File Location: C:Users->NP5048687->eclipse-workspace->JDBCExcel");
				System.out.println("File Name:XSSF-File.xlsx");
				}
				workbook1.close();

				
				System.out.println("<------------------creating another workbook----------------------->");
				
				//hssf workbook
				Workbook workbook2 = new HSSFWorkbook();
				Sheet sheet2 = workbook2.createSheet("Random2");
				
				Font headerFont2 = workbook2.createFont();
				headerFont2.setBold(true);
				headerFont2.setFontHeightInPoints((short)16);
				headerFont2.setColor(IndexedColors.GOLD.getIndex());
				
				CellStyle headerCellStyle2 = workbook2.createCellStyle();
				headerCellStyle2.setFont(headerFont2);
					
				Row headerRow2 = sheet2.createRow(0);
				
				for(int i=0; i < columns.length;i++)
				{
					Cell cell = headerRow2.createCell(i);
					cell.setCellValue(columns[i]);
					cell.setCellStyle(headerCellStyle2);
				}
				
				//cell styling
				CellStyle style2 = workbook2.createCellStyle();
				DataFormat format = workbook2.createDataFormat();
				
				int rowNum2 = 1;
				while(rowNum2 < 65536)
				{
					Row row = sheet2.createRow(rowNum2++);
					
					row.createCell(0).setCellValue(rowNum2-1);
					
				}
				
				
							
				try {
				FileOutputStream fileout2 = new FileOutputStream("HSSF-File.xls");
				workbook2.write(fileout2);
				fileout2.close();
				}
				catch(FileNotFoundException f)
				{
					f.getMessage();
				}
				finally
				{
				System.out.println("Data written successfully...!");
				System.out.println("File Location: C:Users->NP5048687->eclipse-workspace->JDBCExcel");
				System.out.println("File Name: HSSF-File.xlsx");
				}
				workbook2.close();
			}
	
}
