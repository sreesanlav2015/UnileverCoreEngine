package com.cucumber.framework.stepDef;

import org.apache.log4j.Logger;

import com.cucumber.framework.PageObjects.Pulse;
import com.cucumber.framework.TestBase.TestBase;
import com.cucumber.framework.helper.Logger.LoggerHelper;

import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;



public class PulseTest {
	private final Logger log = LoggerHelper.getLogger(PulseTest.class);

	Pulse pulse;
	
	@Given("^Steps for pulse in case of failure$")
	public void steps_for_pulse_in_case_of_failure() throws Throwable {
		  System.out.println("Steps to reproduce:");
		  System.out.println("Open the browser and navigate to the url");
		  System.out.println("I enter username as 'csinboundagent'");
		  System.out.println("I enter password as 'rules'");
		  System.out.println("I click on Login button");
		  System.out.println("Login should be successful");
		  System.out.println("I click on Pulse Link");
		  System.out.println("I enter 'test message' in pulse textbox");
		  System.out.println("I click on post button");
		  System.out.println("'test message' should be posted successfully");
		  System.out.println("When I click on down arrow");
		  System.out.println("I click on Logout");
		  System.out.println("Logout should be successful");
	}
	
	@When("^I click on Pulse Link$")
	public void i_click_on_Pulse_Link() throws Throwable {
	 pulse=new Pulse(TestBase.getDriver());
	 pulse.sendPulseObject(pulse);
	 pulse.clickOnPulseLink();
	
	}

	@When("^I enter \"([^\"]*)\" in pulse textbox$")
	public void i_enter_in_pulse_textbox(String text_msg) throws Throwable {
	  pulse.postText(text_msg);
	  
	}

	@When("^I click on post button$")
	public void i_click_on_post_button() throws Throwable {
	  pulse.clickOnPostBtn();
	 
	}

	@Then("^\"([^\"]*)\" should be posted successfully$")
	public void should_be_posted_successfully(String posted_msg) throws Throwable {
	  pulse.verifyPostedMsg(posted_msg);
	 
	}





}
