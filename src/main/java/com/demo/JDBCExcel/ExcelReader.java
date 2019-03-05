package com.demo.JDBCExcel;


import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReader 
{	
	public static final String filepath = "C:\\Users\\NP5048687\\Documents\\Nikita\\My Data\\Book1.xlsx";
    public static void main(String[] args ) throws IOException, InvalidFormatException,FileNotFoundException
    {
      Workbook workbook = WorkbookFactory.create(new File(filepath));
      int numberOfSheets = workbook.getNumberOfSheets();
      if(numberOfSheets == 0)
      {
    	  System.out.println("File is empty...!");
      }
      else
      {
      System.out.println("Workbook has "+numberOfSheets+" sheets!");
      
      //lambda expression for sheet name
      System.out.println("Retrieving sheets using java 8 lamba");
      workbook.forEach(sheet ->{
    	  System.out.println("=>"+sheet.getSheetName());
    	  
    	  
       //getting the sheet at zeroth index
    	  Sheet sheet1 = workbook.getSheetAt(0);
    	  
       //data formatter for formatting each cell's value as string
    	  DataFormatter df = new DataFormatter();
    	      	  
      //lambda expression for sheet data  
    	  System.out.println("Retrieving sheet data using java 8 lamba");
    	  sheet1.forEach(row -> {
    		  row.forEach(cell -> {
    			  String cellvalue = df.formatCellValue(cell);
    			  System.out.print(cellvalue + "\t");
    		  });
    		  System.out.println();
    	  });
    	  
      });
      
      //closing the workbook
      workbook.close();
      }
    }
}
