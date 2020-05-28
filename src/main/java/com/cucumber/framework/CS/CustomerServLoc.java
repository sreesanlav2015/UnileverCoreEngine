package com.cucumber.framework.CS;

public interface CustomerServLoc {

	String search_txtbox_xpath="//div[@node_name='CPMSearch']//input[@id='pySearchText' and @title='Search for a Case']";
	String search_result_xpath="//tr[@id='$PpySearchResults$ppxResults$l1']";
	String search_results_xpath="//tr[@id='$PpySearchResults$ppxResults$l1']/td";
	String testpath="//i[@name='CPMSearch_pyDisplayHarness_2']";
	void searchFunc(String caseid,String i);
	String addtaskbtn_xpath="//button[@title='Add Task']";
	String addtasksbutton_xpath="//button[@class='Strong pzhc pzbutton' and text()='Add tasks']";
	String wrapupbtn_xpath="//button[@title='Wrap Up']";
	String caseid_xpath="//span[@class='heading_3']";
	String workbasketlink_xpath="//h3[@class='layout-group-item-title' and text()='My workbaskets']";
	String viewqueuefordropdown_xpath="//select[@title='View queue for']";
	String remarks_xpath="//textarea[@id='Remarks']";
	String submitbutton_mid_xpath="//div[@class='pzbtn-mid' and text()='Submit']";
	String submitbutton_xpath="//button[@class='Strong pzhc pzbutton' and text()=' Submit ']";
	String closebtn_mid_xpath ="//div[@class='pzbtn-mid' and text()='Close']";
	String xpathstart="//*[@id='$PpgRepPgSubSectionCPMMyWorkBasketListBB$ppxResults$l";
	String xpathend="']/td[1]/div/span";
}
