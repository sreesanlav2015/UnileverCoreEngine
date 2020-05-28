package com.cucumber.framework.PageObjects;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;

import com.cucumber.framework.CS.CustomerServ;

import cucumber.api.java.Before;

public class NavigateToURL extends CustomerServ {

	public NavigateToURL(WebDriver driver) {
		super(driver);
		
	}
	

	public LoginPage launch() {
	/*driver.get(name);
	driver.manage().window().maximize();
	driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);*/
	return new LoginPage(driver);
	}

	
	

}
