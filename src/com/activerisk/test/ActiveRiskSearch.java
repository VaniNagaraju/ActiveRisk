package com.activerisk.test;

import java.awt.AWTException;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import org.openqa.selenium.firefox.FirefoxDriver;

public class ActiveRiskSearch  
{
	
	public void search(WebDriver driver,String searchTerm) throws InterruptedException
	{
	driver.findElement(By.id("s")).sendKeys(searchTerm,Keys.ENTER);
	Thread.sleep(6000);
	}
	
	public static void main(String[] args) throws InterruptedException, AWTException, IOException
	{
		
		FileInputStream fis=new FileInputStream("./Data/Data.xlsx");
		Workbook workbook=new XSSFWorkbook(fis);
		Sheet sheet1=workbook.getSheetAt(0);
		
		Row r=null;
	    Cell c=null;
	    
		if(sheet1!=null)
		{
			int rowcount=sheet1.getLastRowNum();
			for(int i=1;i<=rowcount;i++)
			{
				r=sheet1.getRow(i);
				if(r!=null)
				{
					int colcount=r.getLastCellNum();
					for(int j=0;j<colcount;j++)
					{
						c=r.getCell(j);
						Cell c1=r.getCell(j+1);
						if(c!=null)
						{
							String searchTerm=r.getCell(j).toString(); // Getting Search term from excel file
							Thread.sleep(1000);

							// launching the browser and visiting sword-activerisk website
							
							System.setProperty("webdriver.gecko.driver","./Drivers/geckodriver.exe");
							WebDriver driver=new FirefoxDriver();
							driver.manage().timeouts().implicitlyWait(10l,TimeUnit.SECONDS);				
							driver.get("http://www.sword-activerisk.com/");
							driver.findElement(By.id("s")).sendKeys(searchTerm,Keys.ENTER);
							Thread.sleep(6000);
							
							//checking whether results found or not
							
							
							try
							{
						     WebElement searchText=driver.findElement(By.xpath("//div[@class='search-results']/h3[text()='Sorry, no results found! Please try again.']"));
							
										if(c1==null)
										{
										   c1=r.createCell(j+1);
									       c1.setCellType(CellType.STRING);
									       c1.setCellValue(searchText.getText());
									       driver.quit();
										}
							}		
							catch(Exception e)
							{
								if(c1==null)
								{
								c1=r.createCell(j+1);
							    c1.setCellType(CellType.STRING);
							    c1.setCellValue("Result Found");
							    
							    List<WebElement> allLinks = driver.findElements(By.xpath("//div[@id='internal-page-content']//a//time"));
								List<WebElement> allLinksurl = driver.findElements(By.xpath("//div[@id='internal-page-content']//span[@class='more-excerpt']/a"));
								int count = allLinks.size();
															
							    if(count>0)
							    {
								for(int k=0;k<count;k++)
								{
								WebElement link = allLinks.get(k);
								WebElement linkInPragraph = allLinksurl.get(k);
								if(link.isDisplayed())
								{
								String text = link.getText();
								String linkUrl = linkInPragraph.getAttribute("href");
								linkInPragraph.sendKeys(Keys.CONTROL,Keys.RETURN);
								
								}
							}
					}
							    
							   	Thread.sleep(4000);
							   	
							   	//Following code is for checking the resulting search links work and load correct pages
							   	
							   	Sheet resultSheet=workbook.getSheet(searchTerm);
							   	
							   	 Set<String> allWindowHandles = driver.getWindowHandles();
							   	 int rv=1;
							                  
							   	 for(String windowHandle:allWindowHandles)
					              {
							               Thread.sleep(3000);
			   	                           String actualTitle = driver.switchTo().window(windowHandle).getTitle();
                  						                							   	    
                  							r=resultSheet.getRow(rv);
                  							String expectedTitle=r.getCell(0).toString();
                  							int cv=1;
                  							c=r.getCell(cv);
                  							c=r.createCell(cv);
        									c.setCellType(CellType.STRING);
							                c.setCellValue(actualTitle);
							                
							                // comparing expected and actual results and writing the result into excel file
							                if(expectedTitle.equals(actualTitle))
							                 {
							                	c1=r.getCell(cv+1);	
							                	c1=r.createCell(cv+1);
							                	c1.setCellValue("pass");
							                 }
							                
							                else
							                	{
							                	  c1=r.getCell(cv+1);	
							                	  c1=r.createCell(cv+1);
							                	  c1.setCellValue("Fail");
							                	 }
							                	    
							                	 rv++;
							   	                 driver.close();
							   	              }            						   	                                    
                                    								
							                }
										Thread.sleep(3000);
				} } } } } }
		
			FileOutputStream fos=new FileOutputStream(new File("./Data/Data.xlsx"));
			workbook.write(fos);
			fis.close();
			fos.close();

		}
					
}
