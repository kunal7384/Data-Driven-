package autoit;

import java.io.IOException;
import java.util.Scanner;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class AutoIt {
	
	WebDriver driver;
	@Test
	
	public void runAutoItTest() throws IOException, InterruptedException
	{
		
		Scanner sc = new Scanner(System.in);
		System.out.println("Enter FirstName");
	String	firstName = sc.nextLine();
	System.out.println("Enter lastName");
	String	lastName = sc.nextLine();
	
	
	 driver = new ChromeDriver();
	 driver.get("http://toolsqa.com/automation-practice-form/");
	 driver.findElement(By.xpath("//*[@name='firstname']")).sendKeys(firstName);
	 driver.findElement(By.xpath("//*[@name='lastname']")).sendKeys(lastName);
	 driver.findElement(By.xpath("//*[@id='sex-0']")).click();
	 driver.findElement(By.xpath("//*[@id='datepicker']")).sendKeys("10/08/2018");
	 driver.findElement(By.xpath("//*[@id='profession-1']")).click();
	 driver.findElement(By.xpath("//*[@id='photo']")).click();
	 Thread.sleep(2000);
	 Runtime.getRuntime().exec("C:\\Users\\dkunal\\Desktop\\Report\\today\\tool.exe");
		System.out.println("photo upload");
	}

}
