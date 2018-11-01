package action.element;

import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

public class ActionandElement {
	
	WebDriver driver;
	@Test(priority=1)
	public void run() throws InterruptedException
	{
		
	driver = new ChromeDriver();
	driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
	driver.manage().timeouts().setScriptTimeout(30, TimeUnit.SECONDS);
	driver.manage().window().maximize();
	driver.get("http://stqatools.com");
	String title = driver.getTitle();
	System.out.println(title);
	
	/*
	// Find All The Element and Perform Click Operation 
	List<WebElement> element = driver.findElements(By.tagName("a"));
	int size_ele=element.size();
	System.out.println("Total Number of Element present on Webpage is ::"+size_ele);
	 String[] elements = new String[size_ele];
	for(int i=0;i<element.size();i++)
	{
		String element_text=element.get(i).getAttribute("href");
		elements[i]=element.get(i).getAttribute("href");
		System.out.println(element_text);
		
		
	}
	for(int i=0;i<element.size();i++)
		{
			
		driver.navigate().to(elements[i]);
			
		}*/
	
	
	WebElement ele = driver.findElement(By.xpath("//*[text()='About']"));
	Thread.sleep(2000);
	Actions act = new Actions(driver);
	Thread.sleep(2000);
	act.moveToElement(ele).build().perform();
	Thread.sleep(2000);
	driver.findElement(By.xpath("//*[text()='Contact Us']")).click();
    
	} 
	
	@Test(priority=2)
    	public void totalqa() throws InterruptedException
    	{
    	driver.get("http://total-qa.com/");	
    	String title =driver.getTitle();
    	System.out.println("title of application is "+title);
    	WebElement element = driver.findElement(By.xpath("//*[text()='Java Basics']"));
    	System.out.println(element.getText());
    	Thread.sleep(2000);
    	Actions action = new Actions(driver);
    	Thread.sleep(2000);
    	action.moveToElement(element).build().perform();
    		System.out.println("================================================END +++++++++++++++++++++++++++++++++");
    	}
     
	}


