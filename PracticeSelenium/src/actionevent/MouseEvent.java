package actionevent;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

public class MouseEvent {
	
	WebDriver driver;
	@Test
	
	public void eventMouse() throws InterruptedException
	{
		
	driver = new ChromeDriver();
/*	driver.get("http://artoftesting.com/");
	WebElement ele = driver.findElement(By.xpath("//*[text()='Product Category']"));
	Actions act = new Actions(driver) ;
	act.moveToElement(ele).build().perform();*/
	
	driver.get("http://artoftesting.com/");
	driver.manage().window().maximize();
	Actions act = new Actions(driver) ;
WebElement sub =driver.findElement(By.xpath("//*[@src='images/selenium.jpg']"));
//act.moveToElement(sub).click().perform();

act.contextClick(sub).click().build().perform();

		
	
	
		
		
	}

}
