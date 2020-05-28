package com.cucumber.framework.PageObjects;

import java.util.List;

import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.testng.Assert;

import com.cucumber.framework.CS.CustomerServ;
import com.cucumber.framework.GeneralHelperSel.SeleniumFunc;
import com.cucumber.framework.helper.Logger.LoggerHelper;

public class Pulse extends CustomerServ implements PulseLoc {
	private final Logger log = LoggerHelper.getLogger(Pulse.class);
	Pulse pulse;
	
	public Pulse(WebDriver driver) {
		super(driver);
	}
	
	public void sendPulseObject(Pulse pulse) {
		this.pulse=pulse;
		//System.out.println("In sendPulseObject method pulse page"+this.pulse);
	}

	public void clickOnPulseLink() {
		//System.out.println("In pulse Func method");
		pulse.goToFrameByName("PegaGadget0Ifr");
		pulse.waitFor(2);
		driver.findElement(By.xpath(pulse_xpath)).click();
	}
	
	public void postText(String text) {
		driver.findElement(By.xpath(say_something_xpath)).sendKeys(text);
	}
	
	public void clickOnPostBtn() {
		driver.findElement(By.xpath(post_button_xpath)).click();
	}
	
public boolean verifyPostedMsg(String postedmsg) {
	List<WebElement> postres=driver.findElements(By.xpath(posted_msg_xpath));
	//System.out.println(postres.size());	
	//System.out.println(postres.get(0).getText());
		try {
		Assert.assertTrue(postres.get(0).getText().contains(postedmsg),"Message is not  posted  successfully");
		//System.out.println("verify message method" +driver);
		return true;
		}catch(Exception e) {
		return false;
		}
		
	}
}
