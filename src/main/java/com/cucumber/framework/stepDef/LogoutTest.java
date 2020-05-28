package com.cucumber.framework.stepDef;



import org.apache.log4j.Logger;

import com.cucumber.framework.PageObjects.LogoutPage;
import com.cucumber.framework.TestBase.TestBase;
import com.cucumber.framework.helper.Logger.LoggerHelper;

import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;


public class LogoutTest {
	private final Logger log = LoggerHelper.getLogger(LogoutTest.class);
	LogoutPage logoutpage;
	
	@Given("^Steps for logout in case of failure$")
	public void steps_for_logout_in_case_of_failure() throws Throwable {
		  System.out.println("Steps to reproduce:");
		  System.out.println("Open the browser and navigate to the url");
		  System.out.println("I enter username as 'csinboundagent'");
		  System.out.println("I enter password as 'rules'");
		  System.out.println("I click on Login button");
		  System.out.println("Login should be successful");
		  System.out.println("When I click on down arrow");
		  System.out.println("I click on Logout");
		  System.out.println("Logout should be successful");
		  
	}
	
	@When("^I click on down arrow$")
	public void i_click_on_down_arrow() throws Throwable {
		logoutpage=new LogoutPage(TestBase.getDriver());
		logoutpage.sendLogoutObject(logoutpage);
		logoutpage.clickOnDownArrow();
	    //System.out.println("Down arrow" +TestBase.getDriver());
	}
	
	@When("^I click on Logout button$")
	public void i_click_on_logout_button() throws Throwable {
		logoutpage=new LogoutPage(TestBase.getDriver());
		logoutpage.sendLogoutObject(logoutpage);
		logoutpage.clickOnLogoutButton();
	    //System.out.println("Down arrow" +TestBase.getDriver());
	}

	@When("^I click on Logout$")
	public void i_click_on_Logout() throws Throwable {
		logoutpage.clickOnLogoutLink();
		//System.out.println("click on logout" +TestBase.getDriver());
	}

	@Then("^Logout should be successful$")
	public void logout_should_be_successful() throws Throwable {
		logoutpage.verifyUsernameField_IsDisplayed();
		//System.out.println("logout success verification" +TestBase.getDriver());
	}



}
