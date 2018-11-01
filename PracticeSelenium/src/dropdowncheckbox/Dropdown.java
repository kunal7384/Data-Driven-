package dropdowncheckbox;

import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

public class Dropdown {
	/*
	WebDriver driver;
	@Test
	public void testDropdown() throws InterruptedException
	{
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
		driver.manage().timeouts().setScriptTimeout(30, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.get("https://www.softwaretestingmaterial.com/sample-webpage-to-automate");
		//WebElement ele = driver.findElement(By.xpath("//*[text()='Checkbox:']"));
		//String text=ele.getText();
	//System.out.println(text);
	   driver.findElement(By.xpath("//*[@value='cbselenium']")).click();
	   
	   driver.findElement(By.xpath("//*[@value='cbqtp']")).click();
		// or
		
	   driver.findElement(By.xpath("//*[@value='radioselenium']")).click();
	   
	   
	   WebElement mul = driver.findElement(By.xpath("//*[@name='multipleselect[]']"));
	    
	   Select sele = new Select(mul);
	   
	   List<WebElement> mule = sele.getOptions();
	 String abc= mul.getText();
	 System.out.println(abc);
	  int mulsize= mule.size();
	  System.out.println("Check box Size:::"+mulsize);
	  //sele.selectByIndex(2);
	  
	  for(int i=0;i<mulsize;i++)
	  {
		  
	mule.get(i) .click();
	
		  
	  }
	//	sele.deselectAll();
		Thread.sleep(3000);

	  
	 WebElement drop= driver.findElement(By.xpath("//*[@name='dropdown']"));
	 Select dropdopdown = new Select(drop);
	 List<WebElement> muldrop = dropdopdown.getOptions();
	 int drop_size=muldrop.size();
	 System.out.println("Drop down size::" +drop_size);
	 for(int j = 0;j<drop_size;j++)
	 {
		 
	String droptext=	muldrop.get(j).getText() ;
		
		System.out.println(droptext); 
		
	 }
	  dropdopdown.selectByVisibleText("Automation Testing");
	  WebElement dob = driver.findElement(By.xpath("//*[@name='bday']"));
	dob.sendKeys("08/02/1990");
	} 
*/

  WebDriver driver;
  @Test
  public void test() throws InterruptedException
  {
	  driver = new ChromeDriver();
	  driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	  driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
	  driver.manage().timeouts().setScriptTimeout(30, TimeUnit.SECONDS);
	  driver.manage().window().maximize();
	  driver.get("https://www.softwaretestingmaterial.com/sample-webpage-to-automate");
	  WebElement checkbox = driver.findElement(By.xpath("//*[text()='Checkbox:']"));
	String checkbox_name=  checkbox.getText();
	System.out.println("Names of the checkbox are :::"+checkbox_name);
	WebElement checkbox_element = driver.findElement(By.xpath("//*[@value='cbselenium']"));
	boolean status =checkbox_element.isDisplayed();
	System.out.println(status);
	if(!checkbox_element.isSelected())
	{
		
		System.out.println("Click on that Elemnet if it is displayed");
		Thread.sleep(2000);
		checkbox_element.click();
		
	}
	else
	{
		
		System.out.println("Element is not Present or Displayed ");
	}
	WebElement checkbox_element1 = driver.findElement(By.xpath("//*[@value='cbqtp']"));
	boolean status1=checkbox_element1.isDisplayed();
	if(status1)
	{
		System.out.println("Element1 displayed");
		checkbox_element1.click();
	}
	else
	{
		System.out.println("Element 1 not Displayed");
		
	}
	
	
	WebElement multiple = driver.findElement(By.xpath("//*[@name='multipleselect[]']"));
	
	String multiText=multiple.getText();
	System.out.println(multiText);
	Select multiple_check = new Select(multiple);
	
	List<WebElement> mul =multiple_check.getOptions();
	//System.out.println("Hiiii"+mul);
	int sss= mul.size();
	System.out.println(sss);
	for(int i=0;i<mul.size();i++)
	{
		
		Thread.sleep(2000);
		mul.get(i).click();
		
		
	}
	
	
	WebElement dropdown = driver.findElement(By.xpath("//*[@name='dropdown']"));
	
	String dtext = dropdown.getText();
	System.out.println("DropDown Text are ::::"+dtext);
	
	Select dropsel = new Select(dropdown);
	
	List<WebElement> drop =dropsel.getOptions();
	
	for(int j =0;j<drop.size();j++)
	{
		drop.get(j).click();
		
		
	}
  }

	
	
	
		
	}
	

	
	
	
