package com.cucumber.framework.PageObjects;

public interface LoginPageLoc {
	
	String username_xpath="//input[@id='txtUserID']";
	String password_xpath="//input[@id='txtPassword']";
	String login_btn_xpath="//span[@class='loginButtonText']";
	String verify_msg_xpath="//span[@data-test-id='2018031408355709381206']"; //span[text()='O2CKatowice']
	String loginclosebtn_xpath="//div[@class='pzbtn-mid' and text()='Close']";
    String error_msg_xpath="//div[@id='error' and text()='                   The information you entered was not recognized.          ']";
}
