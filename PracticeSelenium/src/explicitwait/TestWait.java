package explicitwait;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import net.bytebuddy.description.modifier.SynchronizationState;

public class TestWait {

	WebDriver driver;
  @Test
	public void waitTest() throws InterruptedException
	{
	
	driver = new ChromeDriver();
	driver.get("http://seleniumpractise.blogspot.in/2016/08/how-to-use-explicit-wait-in-selenium.html");
	String title = driver.getTitle();
	System.out.println("Title of the Application is "+title);
	 driver.findElement(By.xpath("//*[text()='Click me to start timer']")).click(); 
	 WebDriverWait wait = new WebDriverWait(driver, 30);
	 WebElement ele = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[text()='WebDriver']")));
	  boolean status = ele.isDisplayed();
	  System.out.println(status);
	  if(status)
	  {
		 System.out.println("Element Display"); 
		  
	  }
	  else
	  {
		  System.out.println("Element Not Displsy"); 
		  
		  
	  }
	}
}