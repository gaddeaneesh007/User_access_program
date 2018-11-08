package com.dell.ofs.useraccess;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Platform;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;

public class DriverandReadexcel {
	XSSFWorkbook workbook;
	XSSFSheet sheetwrite;
	WebDriver driver;
	String currentStatus=null;
	public WebDriver Initialization()
	{
		String dPath = System.getProperty("user.dir")+"/IE/IEDriverServer.exe";
		System.setProperty("webdriver.ie.driver", dPath);
		 DesiredCapabilities capabilities =  DesiredCapabilities.internetExplorer();
		 capabilities.setBrowserName("internet explorer");
		 capabilities.setPlatform(Platform.WINDOWS);
		 capabilities.setVersion("11");
		 capabilities.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
		 capabilities.setCapability("acceptInsecureCerts", true);
		 capabilities.setCapability("acceptSslCerts",true);
		 
		driver = new InternetExplorerDriver(capabilities);
		//return driver;	
	
		//String dPath = System.getProperty("user.dir")+"/Chrome_Driver/chromedriver.exe";
		//System.setProperty("webdriver.chrome.driver", dPath);
		//driver = new InternetExplorerDriver();
		return driver;	
		}
	public void openSite(String regionName, String sourceEnv)
	{
		if(regionName.equalsIgnoreCase("DAO") && sourceEnv.equalsIgnoreCase("GE1"))		
		{
			driver.get("http://auswuofsweb01/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("DAO") && sourceEnv.equalsIgnoreCase("GE2"))		
		{
			driver.get("http://auswuofsweb03/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("DAO") && sourceEnv.equalsIgnoreCase("GE3"))		
		{
			driver.get("http://auswuofsweb02.aus.amer.dell.com/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("DAO") && sourceEnv.equalsIgnoreCase("GE4"))		
		{
			driver.get("http://ausuw4ofsapp01.aus.amer.dell.com/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("EMEA") && sourceEnv.equalsIgnoreCase("GE1"))		
		{
			driver.get("http://ausuwofsweb02.aus.amer.dell.com/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("EMEA") && sourceEnv.equalsIgnoreCase("GE2"))		
		{
			driver.get("http://ausuwofsweb01.aus.amer.dell.com/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("EMEA") && sourceEnv.equalsIgnoreCase("GE3"))		
		{
			driver.get("http://ausuwofsweb03.aus.amer.dell.com/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("EMEA") && sourceEnv.equalsIgnoreCase("GE4"))		
		{
			driver.get("http://ausuw4ofsweb01.aus.amer.dell.com/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("APJ") && sourceEnv.equalsIgnoreCase("GE1"))		
		{
			driver.get("http://u4vmofsapsitap3/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("APJ") && sourceEnv.equalsIgnoreCase("GE2"))		
		{
			driver.get("http://ofssit-apjuat.3dnsqa.dell.com/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("APJ") && sourceEnv.equalsIgnoreCase("GE3"))		
		{
			driver.get("http://u4vmge3appsit01.aus.amer.dell.com/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		else if(regionName.equalsIgnoreCase("APJ") && sourceEnv.equalsIgnoreCase("GE4"))		
		{
			driver.get("http://ausuw4ofsapp03.aus.amer.dell.com/OFS/Admin/UserManager.aspx");
			driver.manage().window().maximize();	
		}
		
	}
	public void getOrderStatusDAO(String orderNumber) throws InterruptedException
	{
		driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_ID_ORDER_NUMBER']")).sendKeys(orderNumber);
		driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_ID_SEARCH']")).click();
		String parent=driver.getWindowHandle();
		Actions action = new Actions(driver);
		action.moveToElement(driver.findElement(By.xpath("//*[@id='ctl00_OFSContent_ListBox']/option[3]"))).doubleClick().build().perform();
		Set<String> windows = driver.getWindowHandles();
		for(String window:windows)
		{
			if(!parent.equals(window))
			{
				driver.switchTo().window(window);
			}
		}
		Thread.sleep(5000);
		currentStatus=driver.findElement(By.xpath("//span[@id='spnofsstatus']")).getText();
		System.out.println(currentStatus);	
		driver.close();
		driver.switchTo().window(parent);
	}
	public void getOrderStatusAPJ(String orderNumber) throws InterruptedException
	{
		driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_ID_ORDER_NUMBER']")).sendKeys(orderNumber);
		driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_ID_SEARCH']")).click();
		String parent=driver.getWindowHandle();
		Actions action = new Actions(driver);
		action.moveToElement(driver.findElement(By.xpath("//*[@id='ctl00_OFSContent_ListBox']/option[3]"))).doubleClick().build().perform();
		Set<String> windows = driver.getWindowHandles();
		for(String window:windows)
		{
			if(!parent.equals(window))
			{
				driver.switchTo().window(window);
			}
		}
		Thread.sleep(5000);
		driver.findElement(By.id("lblOrdDetail")).click();
		currentStatus=driver.findElement(By.xpath("//span[@id='ctl00_OFSContent_lblOFSStatusx']")).getText();
		System.out.println(currentStatus);
		driver.close();
		driver.switchTo().window(parent);
		
	}
	public void getOrderStatusEMEA(String orderNumber) throws InterruptedException
	{
		driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_ID_ORDER_NUMBER']")).sendKeys(orderNumber);
		driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_ID_SEARCH']")).click();
		String parent=driver.getWindowHandle();
		Actions action = new Actions(driver);
		action.moveToElement(driver.findElement(By.xpath("//*[@id='ctl00_OFSContent_ListBox']/option[3]"))).doubleClick().build().perform();
		Set<String> windows = driver.getWindowHandles();
		for(String window:windows)
		{
			if(!parent.equals(window))
			{
				driver.switchTo().window(window);
			}
		}
		Thread.sleep(5000);
		currentStatus=driver.findElement(By.xpath("//span[@id='lblOFSStatus']")).getText();
		System.out.println(currentStatus);
		driver.close();
		driver.switchTo().window(parent);
		
	}
	
	public String openOfsSite(String regionName,String DesEnv,String orderNumber) throws InterruptedException
	{
		
		if(regionName.equalsIgnoreCase("DAO / AMERICAS"))
		{
			
			if(DesEnv.equalsIgnoreCase("GE1"))
			{
				driver.get("http://auswuofsweb01/OFS/Lookup/OrderLookUp.aspx");
				Thread.sleep(2000);	
				getOrderStatusDAO(orderNumber);
				
			}
			else if(DesEnv.equalsIgnoreCase("GE2"))
			{
				driver.get("http://auswuofsweb03/OFS/Lookup/OrderLookUp.aspx");
				Thread.sleep(2000);	
				getOrderStatusDAO(orderNumber);
			}
			else if(DesEnv.equalsIgnoreCase("GE3"))
			{
				driver.get("http://auswuofsweb02.aus.amer.dell.com/OFS/Lookup/OrderLookUp.aspx");
				Thread.sleep(2000);	
				getOrderStatusDAO(orderNumber);
			}
			else if(DesEnv.equalsIgnoreCase("GE4"))
			{
				driver.get("http://ausuw4ofsapp01.aus.amer.dell.com/OFS/LookUp/OrderLookup.aspx");
				Thread.sleep(2000);	
				getOrderStatusDAO(orderNumber);
			}
			else
			{
				System.out.println("Invalid Region");
			}
			
		}
		if(regionName.equalsIgnoreCase("APJ"))
		{
			
			if(DesEnv.equalsIgnoreCase("GE1"))
			{
				driver.get("http://u4vmofsapsitap3/ofs/Lookup/OrderLookUp.aspx");
				Thread.sleep(2000);	
				getOrderStatusAPJ(orderNumber);
			}
			else if(DesEnv.equalsIgnoreCase("GE2"))
			{
				driver.get("http://ofssit-apjuat.3dnsqa.dell.com/OFS/Lookup/OrderLookUp.aspx");
				Thread.sleep(2000);	
				getOrderStatusAPJ(orderNumber);
				
			}
			else if(DesEnv.equalsIgnoreCase("GE3"))
			{
				driver.get("http://u4vmge3appsit01.aus.amer.dell.com/OFS/Lookup/OrderLookUp.aspx");
				Thread.sleep(2000);	
				getOrderStatusAPJ(orderNumber);
			}
			else if(DesEnv.equalsIgnoreCase("GE4"))
			{
				driver.get("http://ausuw4ofsapp03.aus.amer.dell.com/OFS/LookUp/OrderLookup.aspx");
				Thread.sleep(2000);	
				getOrderStatusAPJ(orderNumber);
			}
			else
			{
				System.out.println("Invalid Region");
			}
			
		}
		if(regionName.equalsIgnoreCase("EMEA"))
		{
			
			if(DesEnv.equalsIgnoreCase("GE1"))
			{
				driver.get("http://ausuwofsweb02.aus.amer.dell.com/OFS/Lookup/OrderLookUp.aspx");
				Thread.sleep(2000);	
				getOrderStatusEMEA(orderNumber);
			}
			else if(DesEnv.equalsIgnoreCase("GE2"))
			{
				driver.get("http://ausuwofsweb01.aus.amer.dell.com/OFS/Lookup/orderlookup.aspx");
				Thread.sleep(2000);	
				getOrderStatusEMEA(orderNumber);
			}
			else if(DesEnv.equalsIgnoreCase("GE3"))
			{
				driver.get("http://ausuwofsweb03.aus.amer.dell.com/OFS/Lookup/OrderLookUp.aspx");
				Thread.sleep(2000);	
				getOrderStatusEMEA(orderNumber);
			}
			else if(DesEnv.equalsIgnoreCase("GE4"))
			{
				driver.get("http://ausuw4ofsweb01.aus.amer.dell.com/OFS/LookUp/OrderLookup.aspx");
				Thread.sleep(2000);	
				getOrderStatusEMEA(orderNumber);
			}
			else
			{
				System.out.println("Invalid Region");
			}
			
		}
		return currentStatus;
		
	}

}
