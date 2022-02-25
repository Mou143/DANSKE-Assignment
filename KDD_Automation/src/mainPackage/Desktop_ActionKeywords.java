package mainPackage;

import com.itextpdf.text.BadElementException;
//import org.apache.commons.io.FileUtils;
import com.itextpdf.text.Image;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import javax.swing.JPanel;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Cookie;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.io.FileHandler;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.Select;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.log.SysoCounter;
import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.DefaultListSelectionModel;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.ListSelectionModel;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;


public class Desktop_ActionKeywords 
{
	// All the desktop related reusable codes are available below
	
	public static String element;
	public static ArrayList<String> all_elements=new ArrayList();	
	public String label1[];
	public String Celement;
	public String selected[];
	//public Gridbox gridbox;
		static WebDriver driver;
		String browser;
		JavascriptExecutor jse;
		BufferedImage bi;
		public static String winHandleBefore, splitStringValue, retainSplitStringValue;
		static int splitIntegerValue, retainSplitIntegerValue;
		Actions action;
		

		//Open Browser and run URL Method
		public void open(String browser, String url) throws Exception
		{
			//Run Internet Explorer 
			if(browser.equalsIgnoreCase("IE"))
			{/*
				 * System.setProperty(new Constants().ieDriver, new Constants().ieDriverPath);
				 * DesiredCapabilities cap = DesiredCapabilities.internetExplorer();
				 * cap.setCapability(InternetExplorerDriver.IE_ENSURE_CLEAN_SESSION, true);
				 * cap.setCapability(InternetExplorerDriver.INITIAL_BROWSER_URL, url);
				 * cap.setCapability("ignoreZoomSetting", true); driver = new
				 * InternetExplorerDriver(cap); Thread.sleep(5000);
				 * driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				 */
				
				//driver.get(url);
			}
			
			//Run Google Chrome
			else if(browser.equalsIgnoreCase("Chrome"))
			{
				System.setProperty(new Constants().chromeDriver, new Constants().chromeDriverPath);
				driver = new ChromeDriver();
				driver.manage().window().maximize();
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.get(url);
			}
			
			//Run Mozilla Firefox
			else if(browser.equalsIgnoreCase("Firefox"))
			{
				//System.setProperty(new Constants().firefoxDriver, new Constants().firefoxDriverPath);
				driver = new FirefoxDriver();
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();
				Thread.sleep(5000);
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.get(url);
			}
						
			//No browser found
			else
				System.out.println("Invalid Browser..."+browser);
		}
		
		//Perform action on input tags
		public void input(String name, String type, String data) throws Exception
		{
			//Check for suggestions in textbox
			if(data.equalsIgnoreCase("Arrowdown"))
			{
				if(type.equalsIgnoreCase("id"))
					driver.findElement(By.id(name)).sendKeys(Keys.ARROW_DOWN);
				
				else if(type.equalsIgnoreCase("name"))
					driver.findElement(By.name(name)).sendKeys(Keys.ARROW_DOWN);
				
				else if(type.equalsIgnoreCase("class"))
					driver.findElement(By.className(name)).sendKeys(Keys.ARROW_DOWN);
				
				else if(type.equalsIgnoreCase("css"))
					driver.findElement(By.cssSelector(name)).sendKeys(Keys.ARROW_DOWN);
				
				else if(type.equalsIgnoreCase("link"))
					driver.findElement(By.linkText(name)).sendKeys(Keys.ARROW_DOWN);
				
				else if(type.equalsIgnoreCase("partiallink"))
					driver.findElement(By.partialLinkText(name)).sendKeys(Keys.ARROW_DOWN);
				
				else if(type.equalsIgnoreCase("tagname"))
					driver.findElement(By.tagName(name)).sendKeys(Keys.ARROW_DOWN);
				
				else if(type.equalsIgnoreCase("xpath"))
					driver.findElement(By.xpath(name)).sendKeys(Keys.ARROW_DOWN);
				
				else
					System.out.println("Invalid locator..." + name + "_" + type);
			}
			
			//Press ENTER button while in the textbox
			else if(data.equalsIgnoreCase("enter"))
			{
				if(type.equalsIgnoreCase("id"))
					driver.findElement(By.id(name)).sendKeys(Keys.ENTER);
				
				else if(type.equalsIgnoreCase("name"))
					driver.findElement(By.name(name)).sendKeys(Keys.ENTER);
				
				else if(type.equalsIgnoreCase("class"))
					driver.findElement(By.className(name)).sendKeys(Keys.ENTER);
				
				else if(type.equalsIgnoreCase("css"))
					driver.findElement(By.cssSelector(name)).sendKeys(Keys.ENTER);
				
				else if(type.equalsIgnoreCase("link"))
					driver.findElement(By.linkText(name)).sendKeys(Keys.ENTER);
				
				else if(type.equalsIgnoreCase("partiallink"))
					driver.findElement(By.partialLinkText(name)).sendKeys(Keys.ENTER);
				
				else if(type.equalsIgnoreCase("tagname"))
					driver.findElement(By.tagName(name)).sendKeys(Keys.ENTER);
				
				else if(type.equalsIgnoreCase("xpath"))
					driver.findElement(By.xpath(name)).sendKeys(Keys.ENTER);
				
				else
					System.out.println("Invalid locator..." + name + "_" + type);
			}
			
			//Input the string value captured from application
			else if(data.equalsIgnoreCase("stringValue"))
			{
				if(type.equalsIgnoreCase("id"))
					driver.findElement(By.id(name)).sendKeys(retainSplitStringValue);
				
				else if(type.equalsIgnoreCase("name"))
					driver.findElement(By.name(name)).sendKeys(retainSplitStringValue);
				
				else if(type.equalsIgnoreCase("class"))
					driver.findElement(By.className(name)).sendKeys(retainSplitStringValue);
				
				else if(type.equalsIgnoreCase("css"))
					driver.findElement(By.cssSelector(name)).sendKeys(retainSplitStringValue);
				
				else if(type.equalsIgnoreCase("link"))
					driver.findElement(By.linkText(name)).sendKeys(retainSplitStringValue);
				
				else if(type.equalsIgnoreCase("partiallink"))
					driver.findElement(By.partialLinkText(name)).sendKeys(retainSplitStringValue);
				
				else if(type.equalsIgnoreCase("tagname"))
					driver.findElement(By.tagName(name)).sendKeys(retainSplitStringValue);
				
				else if(type.equalsIgnoreCase("xpath"))
					driver.findElement(By.xpath(name)).sendKeys(retainSplitStringValue);
				
				else
					System.out.println("Invalid locator..." + name + "_" + type);
			}
			
			//Input the integer value captured from application
			else if(data.equalsIgnoreCase("integerValue"))
			{
				if(type.equalsIgnoreCase("id"))
					driver.findElement(By.id(name)).sendKeys(String.valueOf(retainSplitIntegerValue));
				
				else if(type.equalsIgnoreCase("name"))
					driver.findElement(By.name(name)).sendKeys(String.valueOf(retainSplitIntegerValue));
				
				else if(type.equalsIgnoreCase("class"))
					driver.findElement(By.className(name)).sendKeys(String.valueOf(retainSplitIntegerValue));
				
				else if(type.equalsIgnoreCase("css"))
					driver.findElement(By.cssSelector(name)).sendKeys(String.valueOf(retainSplitIntegerValue));
				
				else if(type.equalsIgnoreCase("link"))
					driver.findElement(By.linkText(name)).sendKeys(String.valueOf(retainSplitIntegerValue));
				
				else if(type.equalsIgnoreCase("partiallink"))
					driver.findElement(By.partialLinkText(name)).sendKeys(String.valueOf(retainSplitIntegerValue));
				
				else if(type.equalsIgnoreCase("tagname"))
					driver.findElement(By.tagName(name)).sendKeys(String.valueOf(retainSplitIntegerValue));
				
				else if(type.equalsIgnoreCase("xpath"))
					driver.findElement(By.xpath(name)).sendKeys(String.valueOf(retainSplitIntegerValue));
				
				else
					System.out.println("Invalid locator..." + name + "_" + type);
			}
			
			else
			{
				//Enter text using id locator
				if(type.equalsIgnoreCase("id"))
				{	
		        	Thread.sleep(2000);
		        	driver.findElement(By.id(name)).clear();
					driver.findElement(By.id(name)).sendKeys(data);
				}
					
				//Enter text using name locator
				else if(type.equalsIgnoreCase("name"))
				{
	        		Thread.sleep(2000);
	        		driver.findElement(By.id(name)).clear();
					driver.findElement(By.name(name)).sendKeys(data);
				}

				//Enter text using className locator
				else if(type.equalsIgnoreCase("class"))
				{
					driver.findElement(By.className(name)).clear();
					driver.findElement(By.className(name)).sendKeys(data);
				}

				//Enter text using css locator
				else if(type.equalsIgnoreCase("css"))
				{
					driver.findElement(By.cssSelector(name)).clear();
					driver.findElement(By.cssSelector(name)).sendKeys(data);
				}

				//Enter text using linktext locator
				else if(type.equalsIgnoreCase("link"))
				{
					driver.findElement(By.linkText(name)).clear();
					driver.findElement(By.linkText(name)).sendKeys(data);
				}

				//Enter text using partial link locator
				else if(type.equalsIgnoreCase("partiallink"))
				{
					driver.findElement(By.partialLinkText(name)).clear();
					driver.findElement(By.partialLinkText(name)).sendKeys(data);
				}

				//Enter text using tagname locator
				else if(type.equalsIgnoreCase("tagname"))
				{
					driver.findElement(By.tagName(name)).clear();
					driver.findElement(By.tagName(name)).sendKeys(data);
				}

				//Enter text using xpath locator
				else if(type.equalsIgnoreCase("xpath"))
				{
					
	        		Thread.sleep(2000);
	        		driver.findElement(By.xpath(name)).sendKeys(data);
				}

				//Invalid locator
				else
					System.out.println("Invalid locator..." + name + "_" + type);
			}
		}
		
		//Perform click action
		public void click(String name, String type) throws Exception
		{
			//Using id locator
			if(type.equalsIgnoreCase("id"))
			{
				driver.findElement(By.id(name)).click();
			}
			else if(type.equalsIgnoreCase("id")){
				WebElement ele = driver.findElement(By.xpath(name));
		    	JavascriptExecutor executor = (JavascriptExecutor)driver;
		    	executor.executeScript("arguments[0].click();", ele);
			}
			
			//Using classname locator
			else if(type.equalsIgnoreCase("class"))
				driver.findElement(By.className(name)).click();
					
			//Using csspath locator
			else if(type.equalsIgnoreCase("css"))
				driver.findElement(By.cssSelector(name)).click();
					
			//Using link locator
			else if(type.equalsIgnoreCase("link"))
				driver.findElement(By.linkText(name)).click();
					
			//Using name locator
			else if(type.equalsIgnoreCase("name"))
			{
				driver.findElement(By.name(name)).click();
			}
					
			//Using partial link locator
			else if(type.equalsIgnoreCase("partiallink"))
				driver.findElement(By.partialLinkText(name)).click();
					
			//Using tagname locator
			else if(type.equalsIgnoreCase("tagname"))
				driver.findElement(By.tagName(name)).click();
			else if(type.equalsIgnoreCase("tagname")){
				List<WebElement> allOptions = driver.findElements(By.id("epsi-Catalog-RewardBriefContainer"));
				WebElement allOption=driver.findElement(By.id("epsi-Catalog-RewardBriefContainer"));
				List<WebElement> allOptions1=allOption.findElements(By.tagName("a"));   //Overall product image list
				List<WebElement> allOptions2=allOption.findElements(By.className("ellipsis_text"));	//product name list				 
				 for(WebElement abc:allOptions2)
				 {	
					 Thread.sleep(1000);
					 all_elements.add(abc.getText());  
						 System.out.println(all_elements);				 
				}
			}
			
			
			//Using xpath locator
			else if(type.equalsIgnoreCase("xpath"))
				driver.findElement(By.xpath(name)).click();
					
			else if(type.equalsIgnoreCase("xpath")){
				WebElement ele = driver.findElement(By.xpath(name));
		    	JavascriptExecutor executor = (JavascriptExecutor)driver;
		    	executor.executeScript("arguments[0].click();", ele);
			}
			//Invalid locator
			else
				System.out.println("Invalid locator..." + name + "_" + type);
		
			Thread.sleep(1000);
		}
		
		//Capture screenshot on the page
		public void CaptureScreen(String TC) throws Exception
		{
			String dateName = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
			File src=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			try {
				String destination = System.getProperty("user.dir") + "/TestsScreenshots/"+TC+dateName+".png";
				File finalDestination = new File(destination);
				//FileUtils.copyFile(src, finalDestination);
				File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
				FileHandler.copy(scrFile, finalDestination);
				ExcelUtils.document.add(new Paragraph("   "));
				Image image=Image.getInstance(destination);
				image.setAlignment(Image.LEFT);
				image.scaleToFit(550, 900);
				ExcelUtils.document.add(image);
				ExcelUtils.document.add(new Paragraph("   "));
				ExcelUtils.document.add(new Paragraph("   "));
				Thread.sleep(1000);	
			}
			catch (IOException e) 
			{
				System.out.println(e.getMessage());
			} 
			catch (BadElementException e)
			{
				e.printStackTrace();
			}
		}		
		
		public boolean validateObject(String name, String type, String testStep) throws Exception
		{
			WebElement element;
			jse = (JavascriptExecutor) driver; 
			if(type.equalsIgnoreCase("id") && driver.findElement(By.id(name)).isDisplayed())
			{
				element = driver.findElement(By.id(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				return true;
			}
			
			else if(type.equalsIgnoreCase("class") && driver.findElement(By.className(name)).isDisplayed())
			{
				element = driver.findElement(By.className(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				return true;
			}
					
			else if(type.equalsIgnoreCase("css") && driver.findElement(By.cssSelector(name)).isDisplayed())
			{
				element = driver.findElement(By.cssSelector(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				return true;
			}
					
			else if(type.equalsIgnoreCase("link") && driver.findElement(By.linkText(name)).isDisplayed())
			{	
				element = driver.findElement(By.linkText(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				return true;
			}
						
			else if(type.equalsIgnoreCase("name") && driver.findElement(By.name(name)).isDisplayed())
			{
				element = driver.findElement(By.name(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				return true;
			}
							
			else if(type.equalsIgnoreCase("partiallink") && driver.findElement(By.partialLinkText(name)).isDisplayed())
			{
				element = driver.findElement(By.partialLinkText(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				return true;
			}
							
			else if(type.equalsIgnoreCase("tagname") && driver.findElement(By.tagName(name)).isDisplayed())
			{
				element = driver.findElement(By.tagName(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				return true;
			}
							
			else if(type.equalsIgnoreCase("xpath") && driver.findElement(By.xpath(name)).isDisplayed())
			{
				element = driver.findElement(By.xpath(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				return true;
			}
							
			else
			{
				System.out.println("Invalid locator..." + name + "_" + type);
				return false;
			}
		}
		
		
		public boolean ObjectEnable(String name, String type) throws Exception
		{
			WebElement element;
			jse = (JavascriptExecutor) driver; 
			if(type.equalsIgnoreCase("id") && driver.findElement(By.id(name)).isEnabled())
			{
				element = driver.findElement(By.id(name));
				return true;
			}
			else if(type.equalsIgnoreCase("class") && driver.findElement(By.className(name)).isEnabled())
			{
				element = driver.findElement(By.className(name));
				return true;
			}
			else if(type.equalsIgnoreCase("css") && driver.findElement(By.cssSelector(name)).isEnabled())
			{
				element = driver.findElement(By.cssSelector(name));
				return true;
			}
			else if(type.equalsIgnoreCase("link") && driver.findElement(By.linkText(name)).isEnabled())
			{	
				element = driver.findElement(By.linkText(name));
				return true;
			}
			else if(type.equalsIgnoreCase("name") && driver.findElement(By.name(name)).isEnabled())
			{
				element = driver.findElement(By.name(name));
				return true;
			}
			else if(type.equalsIgnoreCase("partiallink") && driver.findElement(By.partialLinkText(name)).isEnabled())
			{
				element = driver.findElement(By.partialLinkText(name));
				return true;
			}
			else if(type.equalsIgnoreCase("tagname") && driver.findElement(By.tagName(name)).isEnabled())
			{
				element = driver.findElement(By.tagName(name));
				return true;
			}
			else if(type.equalsIgnoreCase("xpath") && driver.findElement(By.xpath(name)).isEnabled())
			{
				element = driver.findElement(By.xpath(name));
				return true;
			}
			else
			{
				System.out.println("Invalid locator..." + name + "_" + type);
				return false;
			}
		}
		
		
		public boolean ObjectDisable(String name, String type) throws Exception
		{
			WebElement element;
			jse = (JavascriptExecutor) driver; 
			if(type.equalsIgnoreCase("id") && !driver.findElement(By.id(name)).isEnabled())
			{
				element = driver.findElement(By.id(name));
				return true;
			}
			else if(type.equalsIgnoreCase("class") && !driver.findElement(By.className(name)).isEnabled())
			{
				element = driver.findElement(By.className(name));
				return true;
			}
			else if(type.equalsIgnoreCase("css") && !driver.findElement(By.cssSelector(name)).isEnabled())
			{
				element = driver.findElement(By.cssSelector(name));
				return true;
			}
			else if(type.equalsIgnoreCase("link") && !driver.findElement(By.linkText(name)).isEnabled())
			{	
				element = driver.findElement(By.linkText(name));
				return true;
			}
			else if(type.equalsIgnoreCase("name") && !driver.findElement(By.name(name)).isEnabled())
			{
				element = driver.findElement(By.name(name));
				return true;
			}
			else if(type.equalsIgnoreCase("partiallink") && !driver.findElement(By.partialLinkText(name)).isEnabled())
			{
				element = driver.findElement(By.partialLinkText(name));
				return true;
			}
			else if(type.equalsIgnoreCase("tagname") && !driver.findElement(By.tagName(name)).isEnabled())
			{
				element = driver.findElement(By.tagName(name));
				return true;
			}
			else if(type.equalsIgnoreCase("xpath") && !driver.findElement(By.xpath(name)).isEnabled())
			{
				element = driver.findElement(By.xpath(name));
				return true;
			}
			else
			{
				System.out.println("Invalid locator..." + name + "_" + type);
				return false;
			}
		}
		
			
		//validate the text
		public boolean validateText(String name, String type, String text) throws Exception
		{
			WebElement element;
			jse = (JavascriptExecutor) driver; 
			if(type.equalsIgnoreCase("id"))
			{
				element = driver.findElement(By.id(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				
				if(element.getText().equals(text))
				{
					return true;
				}
				else
				{
				return false;
				}
			}
			
			else if(type.equalsIgnoreCase("class"))
			{
				element = driver.findElement(By.className(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				if(element.getText().equals(text))
				{
					return true;
				}
				else
				{
				return false;
				}
			}
					
			else if(type.equalsIgnoreCase("css"))
			{
				element = driver.findElement(By.cssSelector(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				if(element.getText().equals(text))
				{
					//capture(element, testStep);
					return true;
				}
				else
				{
				return false;
				}
			}
					
			else if(type.equalsIgnoreCase("link"))
			{	
				element = driver.findElement(By.linkText(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				if(element.getText().equals(text))
				{
					return true;
				}
				else
				{
				return false;
				}
			}
						
			else if(type.equalsIgnoreCase("name"))
			{
				element = driver.findElement(By.name(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				if(element.getText().equals(text))
				{
					return true;
				}
				else
				{
				return false;
				}
			}
							
			else if(type.equalsIgnoreCase("partiallink"))
			{
				element = driver.findElement(By.partialLinkText(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				if(element.getText().equals(text))
				{
					return true;
				}
				else
				{
				return false;
				}
			}
							
			else if(type.equalsIgnoreCase("tagname"))
			{
				element = driver.findElement(By.tagName(name));
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);
				if(element.getText().equals(text))
				{
					return true;
				}
				else
				{
				return false;
				}
			}
							
			else if(type.equalsIgnoreCase("xpath"))
			{
				element = driver.findElement(By.xpath(name));	
				//jse.executeScript("arguments[0].scrollIntoView(true);",element);

				if(element.getText().equals(text))
				{
					return true;
				}
				else
				{
				return false;
				}
			}
							
			else
			{
				System.out.println("Invalid locator..." + name + "_" + type);
				return false;
			}
		}
		
	//Wait for specific time
		public void wait(int seconds) throws Exception
		{
			Thread.sleep(seconds * 1000);
		}

		//Clear cookies in mid of session
		public void clearCookies()
		{
			driver.manage().deleteAllCookies();
		}
		
		//Click the browser back button
		public void backButton() throws Exception
		{
			driver.navigate().back();
			Thread.sleep(2000);
		}
		
		//Selects the value from dropdown by index
		public void select(String name, String type, int data) throws Exception
		{
			Select sel;
			WebElement element;
			
			//Using id locator
			if(type.equalsIgnoreCase("id"))
			{
				element = driver.findElement(By.id(name));
				sel = new Select(element);
				sel.selectByIndex(data);
			}
				
			//Using classname locator
			else if(type.equalsIgnoreCase("class"))
			{
				element = driver.findElement(By.className(name));
				sel = new Select(element);
				sel.selectByIndex(data);
			}

			//Using csspath locator
			else if(type.equalsIgnoreCase("css"))
			{
				element = driver.findElement(By.cssSelector(name));
				sel = new Select(element);
				sel.selectByIndex(data);
			}

			//Using link locator
			else if(type.equalsIgnoreCase("link"))
			{
				element = driver.findElement(By.linkText(name));
				sel = new Select(element);
				sel.selectByIndex(data);
			}

			//Using name locator
			else if(type.equalsIgnoreCase("name"))
			{
				element = driver.findElement(By.name(name));
				sel = new Select(element);
				sel.selectByIndex(data);
			}

			//Using partial link locator
			else if(type.equalsIgnoreCase("partiallink"))
			{
				element = driver.findElement(By.partialLinkText(name));
				sel = new Select(element);
				sel.selectByIndex(data);
			}

			//Using tagname locator
			else if(type.equalsIgnoreCase("tagname"))
			{
				element = driver.findElement(By.tagName(name));
				sel = new Select(element);
				sel.selectByIndex(data);
			}

			//Using xpath locator
			else if(type.equalsIgnoreCase("xpath"))
			{
				element = driver.findElement(By.xpath(name));
				sel = new Select(element);
				sel.selectByIndex(data);
			}

			//Invalid locator
			else
				System.out.println("Invalid locator..." + name + "_" + type);
		}
		
		//Selects the value from dropdown by value
		public void selectByValue(String name, String type, String data) throws Exception
		{
			Select sel;
			WebElement element;

			//Using id locator
			if(type.equalsIgnoreCase("id"))
			{
				element = driver.findElement(By.id(name));
				sel = new Select(element);
				sel.selectByValue(data);
			}

			//Using classname locator
			else if(type.equalsIgnoreCase("class"))
			{
				element = driver.findElement(By.className(name));
				sel = new Select(element);
				sel.selectByValue(data);
			}

			//Using csspath locator
			else if(type.equalsIgnoreCase("css"))
			{
				element = driver.findElement(By.cssSelector(name));
				sel = new Select(element);
				sel.selectByValue(data);
			}

			//Using link locator
			else if(type.equalsIgnoreCase("link"))
			{
				element = driver.findElement(By.linkText(name));
				sel = new Select(element);
				sel.selectByValue(data);
			}

			//Using name locator
			else if(type.equalsIgnoreCase("name"))
			{
				element = driver.findElement(By.name(name));
				sel = new Select(element);
				sel.selectByValue(data);
			}

			//Using partial link locator
			else if(type.equalsIgnoreCase("partiallink"))
			{
				element = driver.findElement(By.partialLinkText(name));
				sel = new Select(element);
				sel.selectByValue(data);
			}

			//Using tagname locator
			else if(type.equalsIgnoreCase("tagname"))
			{
				element = driver.findElement(By.tagName(name));
				sel = new Select(element);
				sel.selectByValue(data);
			}

			//Using xpath locator
			else if(type.equalsIgnoreCase("xpath"))
			{
				element = driver.findElement(By.xpath(name));
				sel = new Select(element);
				sel.selectByValue(data);
			}

			//Invalid locator
			else
				System.out.println("Invalid locator for selectByValue..." + name + "_" + type);
		}
		
		public void selectByVisibleText(String name, String type, String data) throws Exception
		{
			Select sel;
			WebElement element;
			jse = (JavascriptExecutor) driver; 
			//Using id locator
			if(type.equalsIgnoreCase("id"))
			{
				element = driver.findElement(By.id(name));
			//	jse.executeScript("arguments[0].scrollIntoView(true);",element);
			//	Thread.sleep(1000);
				sel = new Select(element);
				sel.selectByVisibleText(data);
				Thread.sleep(1000);
			}

			//Using classname locator
			else if(type.equalsIgnoreCase("class"))
			{
				element = driver.findElement(By.className(name));
			//	jse.executeScript("arguments[0].scrollIntoView(true);",element);
			//	Thread.sleep(1000);
				sel = new Select(element);
				sel.selectByVisibleText(data);
				Thread.sleep(1000);
			}

			//Using csspath locator
			else if(type.equalsIgnoreCase("css"))
			{
				element = driver.findElement(By.cssSelector(name));
			//	jse.executeScript("arguments[0].scrollIntoView(true);",element);
			//	Thread.sleep(1000);
				sel = new Select(element);
				sel.selectByVisibleText(data);
				Thread.sleep(1000);
			}

			//Using link locator
			else if(type.equalsIgnoreCase("link"))
			{
				element = driver.findElement(By.linkText(name));
			//	jse.executeScript("arguments[0].scrollIntoView(true);",element);
			//	Thread.sleep(1000);
				sel = new Select(element);
				sel.selectByVisibleText(data);
				Thread.sleep(1000);
			}

			//Using name locator
			else if(type.equalsIgnoreCase("name"))
			{
				element = driver.findElement(By.name(name));
			//	jse.executeScript("arguments[0].scrollIntoView(true);",element);
			//	Thread.sleep(1000);
				sel = new Select(element);
				sel.selectByVisibleText(data);
				Thread.sleep(1000);
			}

			//Using partial link locator
			else if(type.equalsIgnoreCase("partiallink"))
			{
				element = driver.findElement(By.partialLinkText(name));
			//	jse.executeScript("arguments[0].scrollIntoView(true);",element);
			//	Thread.sleep(1000);
				sel = new Select(element);
				sel.selectByVisibleText(data);
				Thread.sleep(1000);
			}

			//Using tagname locator
			else if(type.equalsIgnoreCase("tagname"))
			{
				element = driver.findElement(By.tagName(name));
			//	jse.executeScript("arguments[0].scrollIntoView(true);",element);
			//	Thread.sleep(1000);
				sel = new Select(element);
				sel.selectByVisibleText(data);
				Thread.sleep(1000);
			}

			//Using xpath locator
			else if(type.equalsIgnoreCase("xpath"))
			{
				element = driver.findElement(By.xpath(name));
			//	jse.executeScript("arguments[0].scrollIntoView(true);",element);
			//	Thread.sleep(1000);
				sel = new Select(element);
				sel.selectByVisibleText(data);
				Thread.sleep(1000);
			}

			//Invalid locator
			else
				System.out.println("Invalid locator for selectByValue..." + name + "_" + type);
		}
		
		
		public void Highlight_Object(String name, String type) throws Exception
		{
			WebElement element;
			jse = (JavascriptExecutor) driver; 
			if(type.equalsIgnoreCase("xpath"))
				{
					element = driver.findElement(By.xpath(name));
					jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='5px solid red'", element);
					Thread.sleep(2000);
					jse.executeScript("arguments[0].style.border='0px solid red'", element);
				}
				else if(type.equalsIgnoreCase("id"))
				{
					element = driver.findElement(By.id(name));
					jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
					jse.executeScript("arguments[0].style.border='0px solid red'", element);
				}
				else if(type.equalsIgnoreCase("name"))
				{
					element = driver.findElement(By.name(name));
					//jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
					jse.executeScript("arguments[0].style.border='0px solid red'", element);
				}
				else if(type.equalsIgnoreCase("class"))
				{
					element = driver.findElement(By.className(name));
					//jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
				}
				else if(type.equalsIgnoreCase("css"))
				{
					element = driver.findElement(By.cssSelector(name));
					//jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
				}
				else if(type.equalsIgnoreCase("link"))
				{
					element = driver.findElement(By.linkText(name));
					//jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
				}
				else if(type.equalsIgnoreCase("partialLink"))
				{
					element = driver.findElement(By.partialLinkText(name));
					//jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
				}
				else if(type.equalsIgnoreCase("tagname"))
				{
					element = driver.findElement(By.tagName(name));
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
				}
				  else
				  {
				   // Select the checkbox
					  System.out.println("Element is not visible");
				   }
		} 
		
		public void scroll_view(String name, String type) throws Exception
		{
			WebElement element;
			jse = (JavascriptExecutor) driver; 
			if(type.equalsIgnoreCase("xpath"))
				{
					element = driver.findElement(By.xpath(name));
					jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='5px solid red'", element);
					Thread.sleep(2000);
					jse.executeScript("arguments[0].style.border='0px solid red'", element);
					
				}
				else if(type.equalsIgnoreCase("id"))
				{
					element = driver.findElement(By.id(name));
					jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
					jse.executeScript("arguments[0].style.border='0px solid red'", element);
				}
				else if(type.equalsIgnoreCase("name"))
				{
					element = driver.findElement(By.name(name));
					jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
					jse.executeScript("arguments[0].style.border='0px solid red'", element);
				}
				else if(type.equalsIgnoreCase("class"))
				{
					element = driver.findElement(By.className(name));
					jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
				}
				else if(type.equalsIgnoreCase("css"))
				{
					element = driver.findElement(By.cssSelector(name));
					jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
				}
				else if(type.equalsIgnoreCase("link"))
				{
					element = driver.findElement(By.linkText(name));
					jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
				}
				else if(type.equalsIgnoreCase("partialLink"))
				{
					element = driver.findElement(By.partialLinkText(name));
					jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
				}
				else if(type.equalsIgnoreCase("tagname"))
				{
					element = driver.findElement(By.tagName(name));
					jse.executeScript("arguments[0].scrollIntoView(true);",element);
					Thread.sleep(1000);
					jse.executeScript("arguments[0].style.border='3px solid red'", element);
					Thread.sleep(2000);
				}
				  else
				  {
				   // Select the checkbox
					  System.out.println("Element is not visible");
				   }
		}
		
		//Expand all transactions in estmt page
		public void expand(String name, String type) throws Exception
		{
//			List<WebElement> elements = new ArrayList<WebElement>();
			WebElement element;
			//Using id locator
			if(type.equalsIgnoreCase("id"))
			{
				while(driver.findElement(By.id(name)).isDisplayed())
				{
					driver.findElement(By.id(name)).click();
				}
			}

			//Using classname locator
			else if(type.equalsIgnoreCase("class"))
			{
				element = driver.findElement(By.className(name));
				((JavascriptExecutor) driver).executeScript("arguments[1].click();", element);
				((JavascriptExecutor) driver).executeScript("arguments[2].click();", element);
			}

			//Using csspath locator
			else if(type.equalsIgnoreCase("css"))
			{
				while(driver.findElement(By.cssSelector(name)).isDisplayed())
					driver.findElement(By.cssSelector(name)).click();
			}

			//Using link locator
			else if(type.equalsIgnoreCase("link"))
			{
				while(driver.findElement(By.linkText(name)).isDisplayed())
					driver.findElement(By.linkText(name)).click();
			}

			//Using name locator
			else if(type.equalsIgnoreCase("name"))
			{
				while(driver.findElement(By.name(name)).isDisplayed())
					driver.findElement(By.name(name)).click();
			}

			//Using partial link locator
			else if(type.equalsIgnoreCase("partiallink"))
			{
				while(driver.findElement(By.partialLinkText(name)).isDisplayed())
					driver.findElement(By.partialLinkText(name)).click();
			}

			//Using tagname locator
			else if(type.equalsIgnoreCase("tagname"))
			{
				while(driver.findElement(By.tagName(name)).isDisplayed())
					driver.findElement(By.tagName(name)).click();
			}

			//Using xpath locator
			else if(type.equalsIgnoreCase("xpath"))
			{
				while(driver.findElement(By.xpath(name)).isDisplayed())
					driver.findElement(By.xpath(name)).click();
			}

			//Invalid locator
			else
				System.out.println("Invalid locator..." + name + "_" + type);

			Thread.sleep(1000);
		}
		
		//Close the browser
		public void closeBrowser()
		{
			driver.close();
		}
		
		//Quits the WebDriver and all browser windows
		public void quitBrowser()
		{
			driver.quit();
		}
		
		//Select Frame using the Frame index
		
		public void selectFrameIndex(int data)
		{
				driver.switchTo().frame(data);
		}
		
		//Mouse Hover on the element
		public void mouseHover(String name, String type) throws Exception
		{
			action = new Actions(driver);
			if(type.equalsIgnoreCase("id"))
				action.moveToElement(driver.findElement(By.id(name))).build().perform();
			
			else if(type.equalsIgnoreCase("name"))
				action.moveToElement(driver.findElement(By.name(name))).build().perform();
			
			else if(type.equalsIgnoreCase("class"))
				action.moveToElement(driver.findElement(By.className(name))).build().perform();
			
			else if(type.equalsIgnoreCase("css"))
				action.moveToElement(driver.findElement(By.cssSelector(name))).build().perform();
			
			else if(type.equalsIgnoreCase("link"))
				action.moveToElement(driver.findElement(By.linkText(name))).build().perform();
			
			else if(type.equalsIgnoreCase("partiallink"))
				action.moveToElement(driver.findElement(By.partialLinkText(name))).build().perform();
			
			else if(type.equalsIgnoreCase("tagname"))
				action.moveToElement(driver.findElement(By.tagName(name))).build().perform();
			
			else if(type.equalsIgnoreCase("xpath"))
				action.moveToElement(driver.findElement(By.xpath(name))).build().perform();
			
			else
				System.out.println("Invalid locator..." + name + "_" + type);
			
		}
		
		//Select the parent frame of the current frame
		public void parentFrame()
		{
			driver.switchTo().parentFrame();
		}
		
		

		//Select the child window
		public void selectWindowChild()
		{
			winHandleBefore = driver.getWindowHandle();
			int i = 0;
			for (String winHandle : driver.getWindowHandles())
			{
				driver.switchTo().window(winHandle);
			}
		}
		
		//Select the child window
		public void closeWindowChild() 
		{
			winHandleBefore = driver.getWindowHandle();
			int i = 0;
			for (String winHandle : driver.getWindowHandles()) {
				driver.switchTo().window(winHandle);
				System.out.println("Winhandle=" + winHandle);
				System.out.println("Count=" + ++i);
				driver.close();
			}
		// switch back to main window using this code
		driver.switchTo().window(winHandleBefore);
		}
		public void ParentWindow()
		{
			winHandleBefore = driver.getWindowHandle();
		}
		
		//Select the parent window
		public void selectWindowParent()
		{
			//winHandleBefore = driver.getWindowHandle();
			driver.switchTo().window(winHandleBefore);
		}

		//Click to accept the popup
		public void acceptPopup()
		{
			Alert alert = driver.switchTo().alert();
			alert.accept();
		}

		//Click to dismiss the popup
		public void dismissPopup()
		{
			Alert alert = driver.switchTo().alert();
			alert.dismiss();
		}
		
		public void PressEnterButton() throws Exception
		{
			Robot rb =new Robot();
			rb.keyPress(KeyEvent.VK_ENTER);
			rb.keyRelease(KeyEvent.VK_ENTER);
		}
		
		public void PressTABButton() throws Exception
		{
			Robot rb =new Robot();
			rb.keyPress(KeyEvent.VK_TAB);
			rb.keyRelease(KeyEvent.VK_TAB);
		}
		
		//Clear cache
		public void clearCache() throws Exception
		{
			driver.get("chrome://settings/clearBrowserData");
			Thread.sleep(2000);
			driver.findElement(By.cssSelector("* /deep/ #clearBrowsingDataConfirm")).click();
			Thread.sleep(5000);
		}

		public void newTab(String URL)
		{
			try
			{
			jse = (JavascriptExecutor)driver;
			jse.executeScript("window.open('"+URL+"')");
			}
			catch(Exception e)
			{
				System.out.println(e.getMessage());
			}
		}

		public void refresh()
		{
			driver.navigate().refresh();
		}

	public String getText(String name, String type)
		{
			WebElement element;
			String OriginalChatName="";
			if(type.equalsIgnoreCase("xpath") && driver.findElement(By.xpath(name)).isDisplayed())
			{
				element = driver.findElement(By.xpath(name));	
				OriginalChatName=element.getText();
			}
			else if(type.equalsIgnoreCase("id") && driver.findElement(By.id(name)).isDisplayed())
			{
				element = driver.findElement(By.id(name));	
				OriginalChatName=element.getText();
			}
			else if(type.equalsIgnoreCase("class") && driver.findElement(By.className(name)).isDisplayed())
			{
				element = driver.findElement(By.className(name));	
				OriginalChatName=element.getText();
			}
			else if(type.equalsIgnoreCase("link") && driver.findElement(By.linkText(name)).isDisplayed())
			{
				element = driver.findElement(By.linkText(name));	
				OriginalChatName=element.getText();
			}
			else if(type.equalsIgnoreCase("name") && driver.findElement(By.name(name)).isDisplayed())
			{
				element = driver.findElement(By.name(name));	
				OriginalChatName=element.getText();
			}
			else
			{
				System.out.println("Invalid locator");
			}
			
			return OriginalChatName;
		}

	public void scrollDown()
	{
		jse = (JavascriptExecutor) driver; 
		jse.executeScript("scroll(0, 250);");
	}

	public void scrollUp() 
	{
		jse = (JavascriptExecutor) driver; 
		jse.executeScript("scroll(0, -250);");
		
	}
	
	public void scrollTop()
	{
		jse = (JavascriptExecutor) driver; 
		jse.executeScript("scroll(0, 0);");
	}
	public void closeFrame(){
		driver.switchTo().defaultContent();
	}		
}
