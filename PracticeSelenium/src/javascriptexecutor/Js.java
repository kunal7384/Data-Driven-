package javascriptexecutor;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.security.Credentials;
import org.testng.annotations.Test;

public class Js {
	
 WebDriver driver;
 
 @Test
 
 public void JavaScript()
 {
	driver = new ChromeDriver();
	driver.get("https://www.app.eisontriplethread.com/");
	driver.manage().window().maximize();
	JavascriptExecutor js = (JavascriptExecutor)driver;
	
	WebElement ele = driver.findElement(By.xpath("//*[text()='Hear from us']"));
	//js.executeScript("arguments[0].click();",ele);

    js.executeScript("return window.innerHeight;");
 
     System.out.println(js.executeScript("return window.innerHeight;"));
    js.executeScript("return window.innerWidth;");
    System.out.println(js.executeScript("return window.innerWidth;"));
	js.executeScript("arguments[0].scrollIntoView();",ele);
	
	
	// driver.close();
 }

}
