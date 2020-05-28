package com.cucumber.framework.PageObjects;

public interface PulseLoc {
	
	String pulse_xpath="//a[@class='Dashboard_Left_Nav' and text()='Pulse']";
	void clickOnPulseLink();
	String say_something_xpath="//textarea[@id='pyMessage']";
	void postText(String text);
	String post_button_xpath="//button[@name='pyPostButton_D_pxRenderPosts_1']";
	void clickOnPostBtn();
	String posted_msg_xpath="//div[contains(@id,'M-')]";
	boolean verifyPostedMsg(String postedmsg);
	void sendPulseObject(Pulse pulse);
}
