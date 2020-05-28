package com.cucumber.framework.CS;

import java.util.List;

import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

import com.cucumber.framework.GeneralHelperSel.SeleniumFunc;
import com.cucumber.framework.PageObjects.LogoutPage;
import com.cucumber.framework.helper.Logger.LoggerHelper;
import com.cucumber.framework.stepDef.PulseTest;

public class CustomerServ extends SeleniumFunc implements CustomerServLoc {
	private final Logger log = LoggerHelper.getLogger(CustomerServ.class);
	
	SeleniumFunc seleniumfunc=new SeleniumFunc(driver);

	public CustomerServ(WebDriver driver) {
		super(driver);
	}
	

	public void testCust() {
		System.out.println("In test customer serv" + driver);
	}

	public void searchFunc(String caseid,String value) {
		//System.out.println("In searchFunc method" +caseid+" "+value+" "+seleniumfunc);
	/*	WebElement searchtxtbox=driver.findElement(By.xpath(testpath));
		String frame=SeleniumFunc.goToFrame(searchtxtbox);
		System.out.println(frame);*/
		seleniumfunc.goToFrameByName("PegaGadget0Ifr");
		SeleniumFunc.waitFor(2);
		driver.findElement(By.xpath(testpath)).click();
		SeleniumFunc.waitFor(2);
		driver.findElement(By.xpath(search_txtbox_xpath)).sendKeys(caseid);
		SeleniumFunc.waitFor(2);
		driver.findElement(By.xpath(search_txtbox_xpath)).sendKeys(Keys.ENTER);
		SeleniumFunc.waitFor(2);
	    int i=Integer.parseInt(value);
		/*WebElement search_res=driver.findElement(By.xpath(search_result_xpath));
		SeleniumFunc.goToFrame(search_res);*/
		//SeleniumFunc.goToFrameByName("PegaGadget0Ifr");
		List<WebElement> searchresults=driver.findElements(By.xpath(search_results_xpath));
		//System.out.println(searchresults.size()+" "+searchresults.get(0).getText());
		searchresults.get(i).click();
		seleniumfunc.switchToDefaultContent();
		
	}
	
	//code to select Add Task button
			public void clickAddTaskButton() {
			/*	WebElement AddTaskBtn=driver.findElement(By.xpath(addtaskbtn_xpath));
				String frame1=goToFrame(AddTaskBtn);
				//System.out.println("Actual frame where element present is"+ frame1);
				waitForElement(AddTaskBtn,2);
				AddTaskBtn.click();*/
				
				try {
					SeleniumFunc.xpath_GenericMethod_Click(addtaskbtn_xpath);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}


		//code to select Wrap Up button
				public void clickWrapUpButton() throws Exception {
					xpath_GenericMethod_Click(wrapupbtn_xpath);
		/*
		 * WebElement WrapUpBtn=driver.findElement(By.xpath(wrapupbtn_xpath));
		 * goToFrame(WrapUpBtn); waitForElement(WrapUpBtn,5); JavascriptExecutor
		 * executor = (JavascriptExecutor)driver;
		 * executor.executeScript("arguments[0].click();",WrapUpBtn);
		 * //WrapUpBtn.click(); waitFor(3);
		 */
				}
			
			//code to select service case based on parameter passed
			public void selectServiceCase(String casetext) throws Exception {
			//	waitFor(2);
			//	driver.findElement(By.xpath("//a[@title='"+casetext+"']")).click();
				xpath_GenericMethod_Click("//a[@title='"+casetext+"']");
				

			}
			
			//code to click on AddTasks button after selecting service case
			public void clickAddTasksButton() {
				/*WebElement AddTasksButton=driver.findElement(By.xpath(addtasksbutton_xpath));
				goToFrame(AddTasksButton);
				waitForElement(AddTasksButton,2);
				AddTasksButton.click();
				waitFor(10);*/
				try {
					SeleniumFunc.xpath_GenericMethod_Click(addtasksbutton_xpath);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}

		public void clickSubmit() {
			/*WebElement SubmitButton=driver.findElement(By.xpath(submitbutton_xpath));
				goToFrame(SubmitButton);
				waitForElement(SubmitButton,2);
				SubmitButton.click();
				waitFor(5);*/
			try {
				SeleniumFunc.xpath_GenericMethod_Click(submitbutton_xpath);
			} catch (Exception e) {
				e.printStackTrace();
			}
			
			}
			
			public String getCaseID() {
				WebElement caseId=driver.findElement(By.xpath(caseid_xpath));
				goToFrame(caseId);
				waitForElement(caseId,2);
				String serviceCaseId1=caseId.getText();
				String serviceCaseId2=serviceCaseId1.substring(1, serviceCaseId1.length()-1);
				return serviceCaseId2;
			}
			
	/*
	 * public void selectWB(String value) { WebElement
	 * workBasketLink=driver.findElement(By.xpath(workbasketlink_xpath));
	 * goToFrame(workBasketLink); waitForElement(workBasketLink,2);
	 * workBasketLink.click(); WebElement
	 * viewQueueForDropdown=driver.findElement(By.xpath(viewqueuefordropdown_xpath))
	 * ; goToFrame(viewQueueForDropdown); waitForElement(viewQueueForDropdown,2);
	 * viewQueueForDropdown.click(); Select s=new Select(viewQueueForDropdown);
	 * s.selectByVisibleText(value); }
	 */
			
			public void selectWB(String value) throws Exception {
				SeleniumFunc.xpath_GenericMethod_Click(workbasketlink_xpath);
				SeleniumFunc.xpath_GenericMethod_selectFromDropdownUsingVisibleTextbyclickingOnDropdown(viewqueuefordropdown_xpath, value);
				}
			
			public  static void clickWBResultsRow(String caseid)
			{
				String xpathstart="//*[@id='$PpgRepPgSubSectionCPMMyWorkBasketListBB$ppxResults$l";
				String xpathend="']/td[1]/div/span";
				for(int i=1;i<=100;i++) {
					WebElement caseID=driver.findElement(By.xpath(xpathstart+i+xpathend));
					try {
					if(caseID.getText().equalsIgnoreCase(caseid)){
						caseID.click();
						break;
					}
					}catch(Exception e) {
						System.out.println("No such element"+e.getMessage());
					}
				}
				driver.switchTo().defaultContent();
			}
			

			public void clickWBResultsRow_CS(String xpath, String caseid) throws Exception {
				String xpathstart="//*[@id='$PpgRepPgSubSectionCPMMyWorkBasketListBB$ppxResults$l";
				String xpathend="']/td[4]/div/span";
				try {

					String frames = goToFrameByTag_NameByXpath(xpath);
					System.out.println("using getattribute name");
					if (frames == null) {
						goToFrameByTag_IdByXpath(xpath);
						System.out.println("using getattribute Id");
					}
				} catch (Exception e) {
					e.getMessage();
				}
				for(int i=1;i<=100;i++) {
					WebElement caseID=driver.findElement(By.xpath(xpathstart+i+xpathend));
					try {
					if(caseID.getText().equalsIgnoreCase(caseid)){
						caseID.click();
						break;
					}
					}catch(Exception e) {
						System.out.println("No such element"+e.getMessage());
					}
				}
				driver.switchTo().defaultContent();
				
			}
			
			public void sendRemarks() {
				WebElement Remarks=driver.findElement(By.xpath(remarks_xpath));
				waitFor(3);
				System.out.println("In Remarks method");
				String frame6=goToFrame(Remarks);
				System.out.println("Actual frame where element present is"+ frame6);
				waitForElement(Remarks,5);
				Remarks.click();
				Remarks.sendKeys("Test Data");
				waitFor(3);
				WebElement SubmitButtonMid=driver.findElement(By.xpath(submitbutton_mid_xpath));
				SubmitButtonMid.click();
			}
			
			public void  closeBtn() {
				/*WebElement closeBtnMid=driver.findElement(By.xpath(closebtn_mid_xpath));
		            goToFrame(closeBtnMid);
					waitForElement(closeBtnMid,2);
					closeBtnMid.click();
					waitFor(2);*/
				try {
					SeleniumFunc.xpath_GenericMethod_Click(closebtn_mid_xpath );
				} catch (Exception e) {
					e.printStackTrace();
				}
				
					}
			
			public void  SubmitBtnMid() {
			/*	WebElement SubmitButtonMid=driver.findElement(By.xpath(submitbutton_mid_xpath));
		        goToFrame(SubmitButtonMid);
				waitForElement(SubmitButtonMid,2);
				SubmitButtonMid.click();
				waitFor(2);*/
				try {
					SeleniumFunc.xpath_GenericMethod_Click(submitbutton_mid_xpath );
				} catch (Exception e) {
					e.printStackTrace();
				}
				
				}


		
		
	
}
