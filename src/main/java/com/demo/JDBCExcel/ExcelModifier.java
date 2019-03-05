package com.demo.JDBCExcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelModifier {
	public static void main(String[] args) throws InvalidFormatException, IOException {
	    Workbook workbook = WorkbookFactory.create(new File("poi-generated-file.xlsx")); 
	    
	    Sheet sheet = workbook.getSheetAt(0);
	    
	    Row row =sheet.getRow(1);
	    
	    Cell cell = row.getCell(2);
	    
	    if(cell == null)
	    {
	    	cell = row.createCell(2);
	    }
	    
	    cell.setCellType(CellType.STRING);
	    cell.setCellValue("Employee Name");
	    
	    FileOutputStream fileout = new FileOutputStream("poi-generated-file.xlsx");
	    workbook.write(fileout);
	    System.out.println("Data Updated successfully..!");
	    System.out.println("File Location: C:Users->NP5048687->eclipse-workspace->JDBCExcel");
	    System.out.println("File Name: poi-generated-file.xlsx");
	    fileout.close();
	    workbook.close();
		
		}
	}
		  
		  
