package com.dell.ofs.useraccess;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;


import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.script.FindFailed;
import org.sikuli.script.Screen;


public class User_Access {
	static WebDriver driver;
	
	public static void main(String[] args) throws InterruptedException, IOException{
		DriverandReadexcel di = new DriverandReadexcel();
		String Smessage = null;
		driver=di.Initialization();
		//WebDriverWait wait = new WebDriverWait(driver, 10000);
		//driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);			
		String path = System.getProperty("user.dir")+"/TestData/userdetails.xlsx";
		
		System.out.println(path);
		
		String ordernumber=null;
		
		File file=new File(path);
		FileInputStream inputStream = new FileInputStream(file);	
		XSSFWorkbook wb=new XSSFWorkbook(inputStream);
		XSSFSheet sheet= wb.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		System.out.println(lastRow);
		String regionName=null,userid=null,Roletype=null,sourceEnv=null,First_Name="",Last_Name="",Email_Address="",Win_Auth_Id=null,Manager=null;
		int id=0;
		FileOutputStream fout = null;
		for(int i=1;i<=lastRow;i++)
		{
			if(i>1)
			{
			driver=di.Initialization();
			}
			regionName=sheet.getRow(i).getCell(0).getStringCellValue();
			sourceEnv=sheet.getRow(i).getCell(1).getStringCellValue();
			First_Name=sheet.getRow(i).getCell(2).getStringCellValue();
			Last_Name = sheet.getRow(i).getCell(3).getStringCellValue();
			Email_Address = sheet.getRow(i).getCell(4).getStringCellValue(); 
			Win_Auth_Id = sheet.getRow(i).getCell(5).getStringCellValue();
			Manager = sheet.getRow(i).getCell(6).getStringCellValue();
			Roletype = sheet.getRow(i).getCell(7).getStringCellValue();
			di.openSite(regionName,sourceEnv);
			Thread.sleep(2000);
			Wait<WebDriver> wait = new FluentWait<WebDriver>(driver)
			.withTimeout(60, TimeUnit.SECONDS)
			.pollingEvery(5, TimeUnit.SECONDS)
			.ignoring(NoSuchElementException.class);
			WebElement Add_New_User= driver.findElement(By.id("ctl00_OFSContent_btnAdd"));
			WebElement search= driver.findElement(By.id("ctl00_OFSContent_txtSearch"));
			search.click();
			//Add_New_User.click();
			new Actions(driver).moveToElement(Add_New_User).build().perform();
			Add_New_User.sendKeys(Keys.SPACE);
			System.out.println("Add new User button Clicked");
			Thread.sleep(5000);
			WebElement LastName= driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_txtLastName']"));
			WebElement Email=driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_txtEmail']"));
			WebElement Auth_Id=driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_txtNTAuthId']"));
			WebElement SelectRole=driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_btnSelectRole']"));
			WebElement CreateUser=driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_btnCreate']"));
			WebElement Cancel=driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_btnCancel']"));
			Select Maneger = new Select(driver.findElement(By.xpath("//select[@id='ctl00_OFSContent_ddManager']")));
			Select Role = new Select(driver.findElement(By.xpath("//select[@id='ctl00_OFSContent_lbRoles']")));
			driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_txtFirstName']")).sendKeys(First_Name);;
			LastName.sendKeys(Last_Name);
			Email.sendKeys(Email_Address);
			Auth_Id.sendKeys(Win_Auth_Id);
			Maneger.selectByVisibleText(Manager);
			if(regionName.equalsIgnoreCase("DAO"))
			{
				if(Roletype.equalsIgnoreCase("Admin"))
				{
					Role.selectByVisibleText("Admin Access");
					SelectRole.sendKeys(Keys.TAB);
					SelectRole.sendKeys(Keys.SPACE);
				}
				else if(Roletype.equalsIgnoreCase("Read only"))
				{
					Role.selectByVisibleText("Test");
					Role.selectByVisibleText("Order Lookup");
					Role.selectByVisibleText("Order Lookup Ext");
					SelectRole.sendKeys(Keys.TAB);
					SelectRole.sendKeys(Keys.SPACE);
				}
				else{
					System.out.println("Please enter currect user Role");
				}
			}
			else if(regionName.equalsIgnoreCase("EMEA"))
			{
				if(Roletype.equalsIgnoreCase("Admin"))
				{
					/*List<WebElement> selectOptions = Role.getOptions();
				    for (WebElement temp : selectOptions) 
				    {
				    	Role.selectByVisibleText(temp.getText());
				    }*/
					Role.selectByVisibleText("Admin Access");
					SelectRole.sendKeys(Keys.TAB);
					SelectRole.sendKeys(Keys.SPACE);
				}
				else if(Roletype.equalsIgnoreCase("Read only"))
				{
				Role.selectByVisibleText("TestRole");
				Role.selectByVisibleText("Order Lookup");
				SelectRole.sendKeys(Keys.TAB);
				SelectRole.sendKeys(Keys.SPACE);
				}
				else{
					System.out.println("Please enter currect user Role");
				}
				
			}
			else if(regionName.equalsIgnoreCase("APJ"))
			{
				if(Roletype.equalsIgnoreCase("Admin"))
				{
					/*List<WebElement> selectOptions = Role.getOptions();
				    for (WebElement temp : selectOptions) 
				    {
				    	Role.selectByVisibleText(temp.getText());
				    }*/
					Role.selectByVisibleText("Admin Access");
					SelectRole.sendKeys(Keys.TAB);
					SelectRole.sendKeys(Keys.SPACE);
				}
				else if(Roletype.equalsIgnoreCase("Read only"))
				{
				Role.selectByVisibleText("APJ_Tester");
				Role.selectByVisibleText("Order Lookup");
				Role.selectByVisibleText("View log histroy");
				SelectRole.sendKeys(Keys.TAB);
				SelectRole.sendKeys(Keys.SPACE);
				}
				else{
					System.out.println("Please enter currect user Role");
				}
			}
			
			SelectRole.click();	
			Thread.sleep(2000);
			CreateUser.sendKeys(Keys.TAB);
			String press = Keys.chord(Keys.SHIFT,Keys.TAB); 
			Cancel.sendKeys(press);
			CreateUser.sendKeys(Keys.ENTER);
			//CreateUser.click();
			Thread.sleep(15000);
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_OFSContent_lblUserId']")));
			if(driver.findElements(By.xpath("//input[@id='ctl00_OFSContent_lblUserId']")).size()>0)
			{
				userid =driver.findElement(By.xpath("//input[@id='ctl00_OFSContent_lblUserId']")).getText();
				sheet.getRow(i).createCell(8).setCellValue(userid);
			
			} 
			else
			{
				sheet.getRow(i).createCell(8).setCellValue("User Not Created");
			}
			fout= new FileOutputStream(file);
			wb.write(fout);
			System.out.println("Workbook closed");
			System.out.println("End of the program");
			driver.quit();
			
		}
		fout.close();
		wb.close();
		
		}

	
		
		
	}


