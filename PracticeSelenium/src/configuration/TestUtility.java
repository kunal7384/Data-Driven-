package configuration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class TestUtility {
	
	WebDriver driver;
	
     @Test
	public void runUtility()
	{
		
	Utility ut = new Utility("F:\\Selemnium Practice code Aug 2018\\PracticeSelenium\\property\\project.properties");
	       driver = new ChromeDriver();
		driver.get(ut.geturl());
		driver.findElement(By.id(ut.userlocator())).sendKeys(ut.getuserName());
		
driver.findElement(By.id(ut.passwordlocator())).sendKeys(ut.getuserpassword());

driver.findElement(By.xpath(ut.clicklocator())).click();
	}

}
