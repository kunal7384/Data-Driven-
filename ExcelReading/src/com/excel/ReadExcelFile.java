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
WebDriver driver;
	
@Test(dataProvider= "dataaa")

public void teststs(String u , String p) throws InterruptedException
 {
	driver = new ChromeDriver();
	
	driver.get("http://demo.guru99.com/V4/");
	
	driver.findElement(By.name("uid")).sendKeys(u);
	
	driver.findElement(By.name("password")).sendKeys(p);
	
	if(isAlertisPresent()==true)
	{
		
	driver.switchTo().alert().accept();
	driver.switchTo().defaultContent();
	System.out.println("Invalid login " + u   + p);
	}
	else
	{
		
	driver.findElement(By.xpath("//a[contains(text(),'Log out')]")).click();
	
	driver.switchTo().alert().accept();
	driver.switchTo().defaultContent();
	System.out.println("valid login " + u   + p);
	Thread.sleep(2000);
		
		
	}
		
		
	}
	
	
 


public boolean isAlertisPresent()
{
	
	try {
		
	driver.switchTo().alert();
	return true;
		
	}catch (NoAlertPresentException e) {
		return false;
	}
	
	
	
}
	
	@DataProvider(name="dataaa")
	public String[][] getexceldata()
	{
		
		System.out.println("Welcome");
		try
		{
			
		
		
	File file = new File("C:\\Users\\dkunal\\Desktop\\test.xls");	
	FileInputStream fis = new FileInputStream(file)	;
//	XSSFWorkbook wb = new XSSFWorkbook(fis);
	
	
	 org.apache.poi.ss.usermodel.Workbook wb = WorkbookFactory.create(fis);
//	XSSFSheet st = wb.getSheetAt(0);
	 
	 org.apache.poi.ss.usermodel.Sheet st = wb.getSheetAt(0);
	int row =  st.getLastRowNum()+1;
	int col = st.getRow(0).getLastCellNum();
	
	String[][] data = new String[row-1][col];
	  
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
			
		}
		return null;
	
	
		
		
	
		
	
	
	
	}
	}

