package com.cucumber.framework.PageObjects;

public interface LogoutPageLoc {
	//a[@title='CS Inbound Officer (Loan, Card)']
	String logout_icon_xpath="//i[@data-test-id='px-opr-image-ctrl']";
    String logout_xpath="//a//span[contains(text(),'Logout')]";
    String username_xpath="//input[@id='txtUserID']";
    String login_popup_close_xpath="//div[@class='pzbtn-mid' and text()='Close']";

}
