package com.cucumber.framework.stepDef;

import com.cucumber.framework.utility.SendEmailUtility;


import cucumber.api.java.en.Given;
public class Sendemail {

	@Given("Enter {string} and {string} and {string} to send an email")
	public void enter_and_and_to_send_an_email(String filePath, String fileName, String sheetName) throws Exception {
		
	
		SendEmailUtility.readExcel(filePath, fileName, sheetName);
	    
	}

} 
