package windowhandles;

import java.util.Iterator;
import java.util.Set;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import net.bytebuddy.matcher.EqualityMatcher;

public class WindowHandle {
	
/*	WebDriver driver;
	
	@Test
	public void handle()
	{
	driver = new ChromeDriver();
	
	driver.get("http://seleniumpractise.blogspot.com/2017/07/multiple-window-examples.html");
	String parent = driver.getWindowHandle();
	System.out.println(parent);
	driver.findElement(By.xpath("//*[text()=' Click this link to start new Tab 'and @href='http://www.google.com'] ")).click();
	Set<String> allwindow=driver.getWindowHandles();
	for(String child:allwindow)
	{
		System.out.println(child);

 
	if(!parent.equalsIgnoreCase(child))
	{
		
		driver.switchTo().window(child)	;
		driver.findElement(By.id("lst-ib")).sendKeys("ABCD");
		driver.findElement(By.name("btnK")).sendKeys(Keys.ENTER);
		
	}
	}
    driver.switchTo().window(parent);
    driver.findElement(By.xpath("//*[text()=' Click this link to start  Session in same window']")).click();
	}*/
	
	
	// trying other way....
	
/*WebDriver driver;
	
	@Test
	public void handle()
	{
	driver = new ChromeDriver();
	
	driver.get("");
	String parent = driver.getWindowHandle();
	System.out.println(parent);
	driver.findElement(By.xpath("//*[text()=' Click this link to start new Tab 'and @href='http://www.google.com'] ")).click();
	Set<String> allwindow=driver.getWindowHandles();
	Iterator<String> it = allwindow.iterator();
	while(it.hasNext())	
	{
		
		String child = it.next();
		
	if(!parent.equalsIgnoreCase(child))
	{
		driver.switchTo().window(child)	;
		driver.findElement(By.id("lst-ib")).sendKeys("ABCD");
		driver.findElement(By.name("btnK")).sendKeys(Keys.ENTER);
	}
	}
	
    driver.switchTo().window(parent);
    driver.findElement(By.xpath("//*[text()=' Click this link to start  Session in same window']")).click();
	}*/
	
	
// FOr practice purpose 
	
/*	WebDriver driver;
	@Test
	
	public void testhandle()
	{
		driver = new ChromeDriver();
		driver.get("http://seleniumpractise.blogspot.com/2017/07/multiple-window-examples.html");
		String title = driver.getTitle();
		System.out.println(title);
	String parent=	driver.getWindowHandle();
	System.out.println("Parent Window is ::"+parent);
	driver.findElement(By.xpath("//*[text()=' Click this link to start new Tab ' and @href='http://www.google.com']")).click();
	Set<String> allwindow = driver.getWindowHandles();
	Iterator<String> it = allwindow.iterator();
	
	while(it.hasNext())
	{
		String child = it.next();
		
	if(!parent.equalsIgnoreCase(child))	
	{
		driver.switchTo().window(child);
	driver.findElement(By.id("lst-ib")).sendKeys("abcd");	
		driver.findElement(By.name("btnK")).sendKeys(Keys.ENTER);
		driver.findElement(By.xpath("//*[@href='https://en.wikipedia.org/wiki/ABCD_2']")).click();
	}


	}
	
	driver.switchTo().window(parent);
		driver.findElement(By.xpath("//*[text()=' Click this link to start  Session in same window'] ")).click();
	
	}*/
	
	
	WebDriver driver;
	
	@Test
	
	public void run()
	{
		
	driver = new ChromeDriver();
	driver.get("http://seleniumpractise.blogspot.com/2017/07/multiple-window-examples.html");
	String title = driver.getTitle();
	System.out.println(title);
	String parent =driver.getWindowHandle();
	System.out.println("Parent Window::" +parent);
	driver.findElement(By.xpath("//*[@href='http://www.google.com']")).click();
	Set<String> allwindow=driver.getWindowHandles();
	System.out.println("All Window" + allwindow);
	
	Iterator<String> it = allwindow.iterator();
	
	
	
	
	while(it.hasNext())
	{
		String child =it.next();
		System.out.println("child Window ::"+child); 
		driver.switchTo().window(child);
		if(!parent.equalsIgnoreCase(child)) {
		driver.findElement(By.id("lst-ib")).sendKeys("Narola Infotech Surat");
		driver.findElement(By.name("btnK")).sendKeys(Keys.ENTER);
		
		}
		
	}
	  driver.switchTo().window(parent);
	  driver.findElement(By.xpath("//*[text()=' Click this link to start  Session in same window']")).click();
	
	}
}
