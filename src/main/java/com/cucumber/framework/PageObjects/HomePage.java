package com.cucumber.framework.PageObjects;

import org.apache.log4j.Logger;
import org.openqa.selenium.WebDriver;

import com.cucumber.framework.CS.CustomerServ;
import com.cucumber.framework.helper.Logger.LoggerHelper;

public class HomePage extends CustomerServ implements HomePageLoc{
	private final Logger log = LoggerHelper.getLogger(HomePage.class);

	public HomePage(WebDriver driver) {
		super(driver);
	}

	public SearchPage search() {
	return new SearchPage(driver);
		
	}

}
