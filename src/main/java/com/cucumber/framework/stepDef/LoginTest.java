package com.cucumber.framework.stepDef;

import org.apache.log4j.Logger;

import com.cucumber.framework.GeneralHelperSel.SeleniumFunc;
import com.cucumber.framework.PageObjects.HomePage;
import com.cucumber.framework.PageObjects.LoginPage;
import com.cucumber.framework.TestBase.TestBase;
import com.cucumber.framework.configreader.ObjectRepo;
import com.cucumber.framework.helper.Logger.LoggerHelper;

import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;














public class LoginTest {
	 private final Logger log = LoggerHelper.getLogger(LoginTest.class);
	
	LoginPage loginpage;
	HomePage homepage;
	
	@Given("^Steps for login in case of failure$")
	public void steps_for_login_in_case_of_failure() throws Throwable {
	  System.out.println("Steps to reproduce:");
	  System.out.println("Open the browser and navigate to the url");
	  System.out.println("I enter username as 'CACSR'");
	  System.out.println("I enter password as 'Rules@1234'");
	  System.out.println("I click on Login button");
	  System.out.println("Login should be successful");
	}

	@Given("^Open the browser and navigate to the url$")
	public void open_the_browser_and_navigate_to_the_url() throws Throwable {
		//System.out.println("Before driver instantiation: "+TestBase.getDriver());
	//	TestBase.getDriver().get(ObjectRepo.reader.getWebsite()); //Run from local
	//	TestBase.getDriver().get(System.getProperty("URL"));  //Run from jenkins
		if(ObjectRepo.reader.getRun().equalsIgnoreCase("jenkins")) {
			TestBase.getDriver().get(System.getProperty("URL"));
		}else {
			TestBase.getDriver().get(ObjectRepo.reader.getWebsite());
		}
		loginpage=new LoginPage(TestBase.getDriver());
		loginpage.sendLoginObject(loginpage);
		//System.out.println("After driver instantiation: " +TestBase.getDriver());
	}

	@When("^I enter username as \"([^\"]*)\"$")
	public void i_enter_username_as(String username) throws Throwable {
		loginpage.setUserName(username);

		//System.out.println("user name is " + username);



	}

	@When("^I enter password as \"([^\"]*)\"$")
	public void i_enter_password_as(String password) throws Throwable {
		loginpage.setPassword(password);
		//System.out.println("Enter password " +password);
	}

	@When("^I click on Login button$")
	public void i_click_on_Login_button() throws Throwable {
	loginpage.clickOnLoginbtn();
	//System.out.println("click on login " +TestBase.getDriver());
	
	}

	@Then("^Login should be successful$")
	public void login_should_be_successful() throws Throwable {
		loginpage.verifyMsg();
	
	}
	
	@Given("^Steps for login with invalid credentials \"([^\"]*)\" and \"([^\"]*)\" in case of failure$")
	public void steps_for_login_with_invalid_credentials_and_in_case_of_failure(String username, String password) throws Throwable {
		  System.out.println("Steps to reproduce:");
		  System.out.println("Open the browser and navigate to the url");
		  System.out.println("I enter username as" +username);
		  System.out.println("I enter password as" +password);
		  System.out.println("I click on Login button");
		//  System.out.println("Login should not be successful");
	}

	@Then("^User should not be logged in$")
	public void user_should_not_be_logged_in() throws Throwable {
	   loginpage.verifyLoginFail();
	}
	


}
