package configuration;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.WebDriver;

public class Utility {

	WebDriver driver;
	File file ;
	FileInputStream fis;
	Properties pro;
	
	public Utility(String filepath)
	{
		try {
			file = new File(filepath);
			fis = new FileInputStream(file);
			pro = new Properties();
			pro.load(fis);
		
		} 
		
		
		catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
			
	public String geturl()
	{
		
		return pro.getProperty("url");
		
		
	}
	
	public String getuserName()
	{
		
		return pro.getProperty("userName");
		
		
	}
	public String getuserpassword()
	{
		
		return pro.getProperty("passWord");
		
		
	}
	
	public String userlocator()
	{
		
		return pro.getProperty("username.id");
		
		
	}
	
	public String passwordlocator()
	{
		
		return pro.getProperty("password.id");
		
		
	}
	
	
	public String clicklocator()
	{
		
		return pro.getProperty("click.xPath");
		
		
	}
}
