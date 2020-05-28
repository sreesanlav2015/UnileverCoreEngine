package com.cucumber.framework.TestBase;

import java.util.concurrent.TimeUnit;

import org.apache.log4j.Logger;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import com.cucumber.framework.configreader.ObjectRepo;
import com.cucumber.framework.configreader.PropertyFileReader;
import com.cucumber.framework.configuration.browser.BrowserType;
import com.cucumber.framework.configuration.browser.ChromeBrowser;
import com.cucumber.framework.configuration.browser.FirefoxBrowser;
import com.cucumber.framework.configuration.browser.HtmlUnitBrowser;
import com.cucumber.framework.configuration.browser.IExploreBrowser;
import com.cucumber.framework.helper.Logger.LoggerHelper;

import cucumber.api.Scenario;
import cucumber.api.java.After;
import cucumber.api.java.AfterStep;
import cucumber.api.java.Before;





public class TestBase {

	private final Logger log = LoggerHelper.getLogger(TestBase.class);
	public static ThreadLocal<WebDriver> dr = new ThreadLocal<WebDriver>();
	private  WebDriver driver=null;

	
	public  WebDriver getBrowserObject(BrowserType bType) throws Exception {
		
		try {
		
			log.info(bType);

			switch (bType) {

			case Chrome:
				ChromeBrowser chrome = ChromeBrowser.class.newInstance();
				return chrome.getChromeDriver(chrome.getChromeCapabilities());

			case Firefox:
				FirefoxBrowser firefox = FirefoxBrowser.class.newInstance();
				return firefox.getFirefoxDriver(firefox.getFirefoxCapabilities());

			case HtmlUnitDriver:
				HtmlUnitBrowser htmlUnit = HtmlUnitBrowser.class.newInstance();
				return htmlUnit.getHtmlUnitDriver(htmlUnit.getHtmlUnitDriverCapabilities());

			case Iexplorer:
				IExploreBrowser iExplore = IExploreBrowser.class.newInstance();
				return iExplore.getIExplorerDriver(iExplore.getIExplorerCapabilities());
			default:
				throw new Exception(" Driver Not Found : " + new PropertyFileReader().getBrowser());
			}
			
		} catch (Exception e) {
			log.equals(e);
			throw e;
		}
		
		
	}
	
	public void setUpDriver(BrowserType bType) throws Exception {
		
		System.out.println("Driver Before instantiation is: "+ driver);
		driver = getBrowserObject(bType);
		setWebDriver(driver);
		System.out.println("Driver After instantiation is: "+ driver);
		log.debug("InitializeWebDrive : " + driver.hashCode());
		driver.manage().timeouts().pageLoadTimeout(ObjectRepo.reader.getPageLoadTimeOut(), TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(ObjectRepo.reader.getImplicitWait(), TimeUnit.SECONDS);
		driver.manage().window().maximize();
	}
	
	
	@Before
	public void before() throws Exception {
		
		ObjectRepo.reader = new PropertyFileReader();
		
		//Run from Jenkins code
	/*	String browser=System.getProperty("Browser");
		setUpDriver(BrowserType.valueOf(browser)); */  //Run from jenkins
		
		//Run from Local code
	//	setUpDriver(ObjectRepo.reader.getBrowser());  //Run from Local
	//	log.info(ObjectRepo.reader.getBrowser());  //Run from Local 
		

if(ObjectRepo.reader.getRun().equalsIgnoreCase("jenkins")){
			System.out.println(ObjectRepo.reader.getRun());
			String browser=System.getProperty("Browser");
			setUpDriver(BrowserType.valueOf(browser));
		} else {
			System.out.println(ObjectRepo.reader.getRun());
			setUpDriver(ObjectRepo.reader.getBrowser());  
			log.info(ObjectRepo.reader.getBrowser());
		}
	}

	
	@After
	public void after(Scenario scenario) throws Exception {
		 if (scenario.isFailed()) {
	         scenario.embed(((TakesScreenshot)getDriver()).getScreenshotAs(OutputType.BYTES), "image/png");
	        }
		 System.out.println("Driver in After Method is: "+ driver);
		 driver.quit();
	     log.info("Browser closed");
	}
	
	
	public static WebDriver getDriver() {
        return dr.get();
    }
 
    public static void setWebDriver(WebDriver driver) {
        dr.set(driver);
    }
    
    @AfterStep
    public void afterStep(Scenario scenario) {
    	 scenario.embed(((TakesScreenshot)getDriver()).getScreenshotAs(OutputType.BYTES), "image/png");
    }
	
}
