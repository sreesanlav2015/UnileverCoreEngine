package com.cucumber.framework.stepDef;

import java.util.List;

import org.apache.log4j.Logger;

import com.cucumber.framework.PageObjects.SearchPage;
import com.cucumber.framework.TestBase.TestBase;
import com.cucumber.framework.helper.Logger.LoggerHelper;

import io.cucumber.datatable.DataTable;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;




public class SearchTest {
	private final Logger log = LoggerHelper.getLogger(SearchTest.class);

	SearchPage searchpage;
	
	@Given("^Steps for search in case of failure$")
	public void steps_for_search_in_case_of_failure() throws Throwable {
		  System.out.println("Steps to reproduce:");
		  System.out.println("Open the browser and navigate to the url");
		  System.out.println("I enter username as 'csinboundagent'");
		  System.out.println("I enter password as 'rules'");
		  System.out.println("I click on Login button");
		  System.out.println("Login should be successful");
		  System.out.println("I search for the valid caseid");
		  System.out.println("case should be displayed");
		  System.out.println("I click on down arrow");
		  System.out.println("I click on Logout");
		  System.out.println("Logout should be successful");
	}
	
	@When("^I search for the valid caseid$")
	public void i_search_for_the_valid_caseid(DataTable table) throws Throwable {
		searchpage=new SearchPage(TestBase.getDriver());
		//System.out.println(searchpage);
		searchpage.sendSearchObject(searchpage);
		//System.out.println(searchpage);
		List<List<String>> data=table.asLists();
		searchpage.searchCase(data.get(1).get(0).toString(),data.get(1).get(1) );
		//System.out.println(searchpage);
		
	}

	@Then("^case should be displayed$")
	public void case_should_be_displayed() throws Throwable {
	 System.out.println("-----------");
	}



}
