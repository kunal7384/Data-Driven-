package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.lang.invoke.SwitchPoint;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ReadExcelFile {

	public String[][] getexceldata(String path ,String ss)
	{
		
		System.out.println("Welcome");
		try
		{
			
		String data[][] = null;	
		
	File file = new File(path);	
	FileInputStream fis = new FileInputStream(file)	;
//	XSSFWorkbook wb = new XSSFWorkbook(fis);
	
	
	 org.apache.poi.ss.usermodel.Workbook wb = WorkbookFactory.create(fis);
//	XSSFSheet st = wb.getSheetAt(0);
	 
	 org.apache.poi.ss.usermodel.Sheet st = wb.getSheet(ss);
	int row =  st.getLastRowNum()+1;
	int col = st.getRow(0).getLastCellNum();
	
	data = new String[row-1][col];
	  
		Iterator<Row> rowiterator = st.iterator();
		int i=0;
		int t= 0;
	    while(rowiterator.hasNext())
	    {
	    Row rows = rowiterator.next();	
	    	
	    if(i++!=0)
	    {
	    	
	    int k = t;	
	    	
	    t++;	
	    	
	    Iterator<Cell> colsiterator=rows.iterator();
	    int j=0;
	    while(colsiterator.hasNext())
	    {
	    Cell cell = colsiterator.next();
	    		
	    	
	    switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
	    			
	    	System.out.println(k+",");	
	    	System.out.println(j+",");
	    data[k][j++]	=cell.getStringCellValue();
	    	System.out.println(cell.getNumericCellValue());
	    	break;
	    		
	    		
	    	
	  		case Cell.CELL_TYPE_STRING:
	  	    			
	  	    	System.out.println(k+",");	
	  	    	System.out.println(j+",");
	  	    data[k][j++]	=cell.getStringCellValue();
	  	    	System.out.println(cell.getStringCellValue());
	  	    	break;	
	    		
	  		case Cell.CELL_TYPE_BOOLEAN:
				System.out.print(k+",");
				System.out.print(j+",");
				data[k][j++] = cell.getStringCellValue();
				System.out.println(cell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				System.out.print(k+",");
				System.out.print(j+",");
				data[k][j++] = cell.getStringCellValue();
				System.out.println(cell.getStringCellValue());
				break;
	    		
	    	  } 	
	    	
	    }
	    	
	    	System.out.println(",");
	    	
	    }
	    
	    
	    }
	    
	    fis.close();
		return data;
	    }  
		
		catch (Exception e) {
			// TODO: handle exception
		}
		return null;
		
	}
	public static void main(String[] args) {
		
		String path= "C:\\Users\\dkunal\\Desktop\\test.xls";
		ReadExcelFile r = new ReadExcelFile();
		r.getexceldata(path, "Sheet1");
		
	}
	}

