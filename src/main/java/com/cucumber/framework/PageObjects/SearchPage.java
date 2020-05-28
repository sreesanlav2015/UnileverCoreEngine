package com.cucumber.framework.PageObjects;

import org.apache.log4j.Logger;
import org.openqa.selenium.WebDriver;

import com.cucumber.framework.CS.CustomerServ;
import com.cucumber.framework.helper.Logger.LoggerHelper;

public class SearchPage extends CustomerServ implements SearchPageLoc {
	private final Logger log = LoggerHelper.getLogger(SearchPage.class);
	SearchPage searchpage;

	public SearchPage(WebDriver driver) {
		super(driver);
	}

	public void sendSearchObject(SearchPage searchpage) {
		this.searchpage=searchpage;
		//System.out.println("In sendSearchObject method search page"+this.searchpage);
	}
	
	public void searchCase(String caseid,String i) {
		//System.out.println("In searchCase method of search  page1" +searchpage);
		searchpage.searchFunc(caseid, i);
		//System.out.println("In searchCase method of search   page2" +searchpage);
	}


}
