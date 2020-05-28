package com.cucumber.framework.GeneralHelperSel;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.Reporter;

import com.cucumber.framework.CS.CustomerServ;
import com.cucumber.framework.helper.Logger.LoggerHelper;

public class SeleniumFunc implements SeleniumFuncLoc {
	private final Logger log = LoggerHelper.getLogger(SeleniumFunc.class);
	protected static WebDriver driver;
	// public static WebDriver driver;

	public SeleniumFunc(WebDriver driver) {
		this.driver = driver;
	}

	public void testSel() {
		System.out.println("In testSel of SeleniumFunc" + driver);
	}

	/* Method to handle windows based pop-ups */
	public static void robotClass() {
		Robot robot;
		try {
			robot = new Robot();
			robot.keyPress(KeyEvent.VK_TAB);
			robot.keyRelease(KeyEvent.VK_TAB);
			robot.keyPress(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);
		} catch (AWTException e) {
			e.printStackTrace();
		}
		//waitFor(1);
	}

	/* Method to get the current date time */
	public static String getCurrentDateTime() {

		DateFormat dateFormat = new SimpleDateFormat("_yyyy-MM-dd_HH-mm-ss");
		Calendar cal = Calendar.getInstance();
		String time = "" + dateFormat.format(cal.getTime());
		return time;
	}

	/* Method to get the current date only */
	public static String getCurrentDate() {
		return getCurrentDateTime().substring(0, 11);
	}

	/* Method to capture screenshot */
	public String captureScreen(String name) {
		if (name == "") {
			name = "blank";
		}

		File destFile = null;
		Calendar calendar = Calendar.getInstance();
		SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-YYYY-hh_mm_ss");
		File srcFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try {
			String reportDirectory = new File(System.getProperty("user.dir")).getAbsolutePath()
					+ "target\\screenshot\\";
			destFile = new File((String) reportDirectory + name + "_" + formatter.format(calendar.getTime()) + ".png");
			FileUtils.copyFile(srcFile, destFile);

			Reporter.log("<a href='" + destFile.getAbsolutePath() + "'> <img src='" + destFile.getAbsolutePath()
					+ "height='200' width='100'/> </a>");
		} catch (IOException e) {
			e.printStackTrace();
		}
		return destFile.toString();
	}

	/* Method to get the number of frames in the page using Javascript Executor */
	public void getNumberOfFrames() {
		JavascriptExecutor exe = (JavascriptExecutor) driver;
		Integer numberOfFrames = Integer.parseInt(exe.executeScript("return window.length").toString());
		System.out.println("Number of iframes on the page are " + numberOfFrames);
	}

	/* Method to handle stale element reference */
	public static void HandleStaleElement(WebElement element) {
		for (int i = 0; i < 15; i++) {
			try {
				element.click();
				break;
			} catch (StaleElementReferenceException e) {
				e.toString();
				System.out.println("Trying to recover from a stale element :" + e.getMessage());
			}
		}
	}

	public static void HandleStaleElement(String xpath) {
		for (int i = 0; i < 15; i++) {
			try {
				driver.findElement(By.xpath(xpath)).click();
				break;
			} catch (StaleElementReferenceException e) {
				e.toString();
				System.out.println("Trying to recover from a stale element :" + e.getMessage());
			}
		}
	}

	/* Method to refresh the page */
	public void getRefresh() {
		driver.navigate().refresh();
	}

	public String goToFrame(WebElement element) {
		String frameNameone = null;
		List<WebElement> frames = driver.findElements(By.tagName("iframe"));
		System.out.println(frames.size());
		for (int i = 0; i < frames.size(); i++) {
			System.out.println(frames.get(i).getAttribute("id"));
			String frameNametwo = frames.get(i).getAttribute("id");
			driver.switchTo().defaultContent();
			waitFor(5);
			driver.switchTo().frame(frameNametwo);
			try {
				if (element.isDisplayed()) {
					System.out.println("Element" + element + "is present in the frame" + frameNametwo);
					frameNameone = frameNametwo;
					break;
				}
			} catch (Exception e) {
				System.out.println("Element is not present inside the frame" + element);
				driver.switchTo().defaultContent();
			}
		}
		return frameNameone;
	}

	/* Method to switch to frame using frame name or id */
	public void goToFrameByName(String framename) {
		driver.switchTo().frame(framename);
	}

	/* Method to switch out of the frame */
	public void switchToDefaultContent() {
		driver.switchTo().defaultContent();
	}

	/* Method to print the number of windows present */
	public void printAllWindowIds() {
		Set<String> windowids1 = driver.getWindowHandles();
		Iterator<String> iter1 = windowids1.iterator();
		while (iter1.hasNext()) {
			System.out.println(iter1.next());
		}
	}

	/* Method to switch to second window */
	public void switchToSecondWindow() {
		Set<String> windowids1 = driver.getWindowHandles();
		Iterator<String> iter1 = windowids1.iterator();
		/*
		 * while(iter1.hasNext()) { System.out.println(iter1.next()); }
		 */

		String pwindow = iter1.next();
		System.out.println("parent window id is:" + pwindow);
		String cwindow = iter1.next();
		System.out.println("child window id is:" + cwindow);
		waitFor(5);
		driver.switchTo().window(cwindow);
	}

	/* Method to switch to third window */
	public void switchToThirdWindow() {
		Set<String> windowids1 = driver.getWindowHandles();
		Iterator<String> iter1 = windowids1.iterator();
		/*
		 * while(iter1.hasNext()) { System.out.println(iter1.next()); }
		 */

		String pwindow = iter1.next();
		System.out.println("parent window id is:" + pwindow);
		String fcwindow = iter1.next();
		System.out.println("first child window id is:" + fcwindow);
		String scwindow = iter1.next();
		System.out.println("second child window id is:" + scwindow);
		waitFor(5);
		driver.switchTo().window(scwindow);
	}

	/* Method to verify the title of the page */
	public void verifyPageTitle(String value) {
		waitFor(5);
		System.out.println("Title is :" + value);
		Assert.assertTrue(driver.getTitle().contains(value), "Expected page " + value + " not displayed");
	}

	/* Method to select value from drop down using visible text */
	public void selectFromDropdownUsingVisibleText(String expath, String valuetoselect) {
		Select s = new Select(driver.findElement(By.xpath(expath)));
		s.selectByVisibleText(valuetoselect);
	}

	// select value from dropdown using index
	public void selectFromDropdownUsingIndex(String expath, int indexvalue) {
		Select s = new Select(driver.findElement(By.xpath(expath)));
		s.selectByIndex(indexvalue);
	}

	// select value from dropdown using value of the value
	public void selectFromDropdownUsingValue(String expath, String value) {
		Select s = new Select(driver.findElement(By.xpath(expath)));
		s.selectByValue(value);
	}

	/*
	 * Method to click on element based on index value if multiple elements are
	 * found with same locator
	 */
	public void clickOnSelectedElement(String expath, int i) {
		List<WebElement> elements = driver.findElements(By.xpath(expath));
		System.out.println(elements.size());
		clickUsingAction(elements.get(i));
	}

	/* Method to get text of the selected element */
	public String getSelectedElementText(String expath, int i) {
		List<WebElement> elements = driver.findElements(By.xpath(expath));
		System.out.println(elements.size());
		String text = elements.get(i).getText();
		return text;
	}

	/* Method to click on an element using Actions class */
	public void clickUsingAction(WebElement element) {
		Actions action = new Actions(driver);
		action.moveToElement(element).click().build().perform();
	}

	public void clickLinkInTable(String expath, int row, int col, int link) {
		List<WebElement> rows = driver.findElements(By.xpath(expath));
		System.out.println("rows size are: " + rows.size());
		List<WebElement> cells = rows.get(row).findElements(By.tagName("td"));
		System.out.println("cell size are: " + cells.size());
		String linktext = cells.get(col).getText();
		System.out.println("Id is: " + linktext);

		List<WebElement> links = cells.get(col).findElements(By.tagName("a"));
		System.out.println("Link are :" + links.size());
		System.out.println(links.get(link).getText());

		HandleStaleElement(links.get(link));

		System.out.println(links.get(link).getAttribute("href"));
		waitFor(5);

	}

	public void clickLinkInTableDynamically(String expath, String Headertext, int link) {
		List<WebElement> rows = driver.findElements(By.xpath(expath));
		System.out.println("rows size are: " + rows.size());
		List<WebElement> headers = rows.get(0).findElements(By.tagName("th"));
		System.out.println("Header size are: " + headers.size());
		for (int j = 0; j < headers.size(); j++) {
			String headertext = headers.get(j).getText();
			System.out.println("Header is: " + headertext);
			if (headertext.equalsIgnoreCase(Headertext)) {
				System.out.println("Header is: " + headertext + " " + Headertext);
				System.out.println("In Rows for loop");
				for (int i = 1; i < rows.size(); i++) {
					List<WebElement> cells = rows.get(i).findElements(By.tagName("td"));
					System.out.println("cell size are: " + cells.size());
					String linktext = cells.get(j).getText();
					System.out.println("Text is: " + linktext);
					List<WebElement> links = cells.get(j).findElements(By.tagName("a"));
					System.out.println("Link are :" + links.size());
					System.out.println(links.get(link).getText());
					HandleStaleElement(links.get(link));
					System.out.println(links.get(link).getAttribute("href"));
					waitFor(5);
					break;

				}
				break;
			}
		}

	}

	public String getLinkTextFromTableDynamically(String expath, String Headertext, int link) {
		String text = null;
		List<WebElement> rows = driver.findElements(By.xpath(expath));
		System.out.println("rows size are: " + rows.size());
		List<WebElement> headers = rows.get(0).findElements(By.tagName("th"));
		System.out.println("Header size are: " + headers.size());
		for (int j = 0; j < headers.size(); j++) {
			String headertext = headers.get(j).getText();
			System.out.println("Header is: " + headertext);
			if (headertext.equalsIgnoreCase(Headertext)) {
				System.out.println("Header is: " + headertext + " " + Headertext);
				System.out.println("In Rows for loop");
				for (int i = 1; i < rows.size(); i++) {
					List<WebElement> cells = rows.get(i).findElements(By.tagName("td"));
					System.out.println("cell size are: " + cells.size());
					String linktext = cells.get(j).getText();
					System.out.println("Text is: " + linktext);
					List<WebElement> links = cells.get(j).findElements(By.tagName("a"));
					System.out.println("Link are :" + links.size());
					System.out.println(links.get(link).getText());
					text = links.get(link).getText();
					// HandleStaleElement(links.get(link));
					// System.out.println(links.get(link).getAttribute("href"));
					waitFor(5);
					break;

				}
				break;
			}
		}
		return text;

	}

	public void clickButtonInTable(String expath, int row, int col, int button) {
		List<WebElement> rows = driver.findElements(By.xpath(expath));
		System.out.println("rows size are: " + rows.size());
		List<WebElement> cells = rows.get(row).findElements(By.tagName("td"));
		System.out.println("cell size are: " + cells.size());
		String linktext = cells.get(col).getText();
		System.out.println("Id is: " + linktext);
		List<WebElement> buttons = cells.get(col).findElements(By.tagName("button"));
		System.out.println("Buttons are :" + buttons.size());
		System.out.println(buttons.get(button).getText());
		HandleStaleElement(buttons.get(button));
		waitFor(5);

	}

	/* Method to write data to a file */
	public void writeDataToFile(String text) {

		try {
			File file = new File("test.data");
			// Delete old file if exists
			file.delete();
			file.createNewFile();
			FileWriter fileWrite = new FileWriter(file);
			BufferedWriter Bwrite = new BufferedWriter(fileWrite);
			Bwrite.write(text);
			Bwrite.newLine();
			Bwrite.close();
			fileWrite.close();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	/* Method to read data from a file */
	public String ReadDataFromFile() throws IOException {
		String str = null;
		try {
			File file = new File("test.data");
			FileReader fileReader = new FileReader(file);
			BufferedReader Buffreader = new BufferedReader(fileReader);
			String strline;
			while ((strline = Buffreader.readLine()) != null) {
				System.out.println(strline);
				str = strline;
			}
			Buffreader.close();
			fileReader.close();

		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return str;

	}

	/* Method to switch to a frame using tag id and xpath locator */
	public static String goToFrameByTag_IdByXpath(String xpath) throws InterruptedException {
		System.out.println("In GoToFrames Method Start");
		String frameNameone = null;
		List<WebElement> frames = driver.findElements(By.tagName("iframe"));
		System.out.println(frames.size());
		for (int i = 0; i < frames.size(); i++) {
			System.out.println(frames.get(i).getAttribute("id"));
			String frameNametwo = frames.get(i).getAttribute("id");
			driver.switchTo().defaultContent();
			Thread.sleep(1000);
			driver.switchTo().frame(frameNametwo);
			try {
				if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
					System.out.println("Element" + xpath + "is present in the frame :" + frameNametwo);
					frameNameone = frameNametwo;
					break;
				}
			} catch (Exception e) {
				System.out.println("Element is not present inside the frame :" + xpath);
				driver.switchTo().defaultContent();
			}
		}
		System.out.println("In GoToFrames Method End");
		return frameNameone;
	}

	/* Method to switch to a frame using tag id and name locator */
	public static String goToFrameByTag_IdByName(String name) throws InterruptedException {
		System.out.println("In GoToFrames Method Start");
		String frameNameone = null;
		List<WebElement> frames = driver.findElements(By.tagName("iframe"));
		System.out.println(frames.size());
		for (int i = 0; i < frames.size(); i++) {
			System.out.println(frames.get(i).getAttribute("id"));
			String frameNametwo = frames.get(i).getAttribute("id");
			driver.switchTo().defaultContent();
			Thread.sleep(1000);
			driver.switchTo().frame(frameNametwo);
			try {
				if (driver.findElement(By.name(name)).isDisplayed()) {
					System.out.println("Element" + name + "is present in the frame :" + frameNametwo);
					frameNameone = frameNametwo;
					break;
				}
			} catch (Exception e) {
				System.out.println("Element is not present inside the frame :" + name);
				driver.switchTo().defaultContent();
			}
		}
		System.out.println("In GoToFrames Method End");
		return frameNameone;
	}

	/* Method to switch to a frame using tag id and id locator */
	public static String goToFrameByTag_IdById(String id) throws InterruptedException {
		System.out.println("In GoToFrames Method Start");
		String frameNameone = null;
		List<WebElement> frames = driver.findElements(By.tagName("iframe"));
		System.out.println(frames.size());
		for (int i = 0; i < frames.size(); i++) {
			System.out.println(frames.get(i).getAttribute("id"));
			String frameNametwo = frames.get(i).getAttribute("id");
			driver.switchTo().defaultContent();
			Thread.sleep(1000);
			driver.switchTo().frame(frameNametwo);
			try {
				if (driver.findElement(By.id(id)).isDisplayed()) {
					System.out.println("Element" + id + "is present in the frame :" + frameNametwo);
					frameNameone = frameNametwo;
					break;
				}
			} catch (Exception e) {
				System.out.println("Element is not present inside the frame :" + id);
				driver.switchTo().defaultContent();
			}
		}
		System.out.println("In GoToFrames Method End");
		return frameNameone;
	}

	/* Method to switch to a frame using tag name and xpath locator */
	public static String goToFrameByTag_NameByXpath(String xpath) throws InterruptedException {
		System.out.println("In GoToFrames Method Start");

		String frameNameone = null;
		List<WebElement> frames = driver.findElements(By.tagName("iframe"));
		System.out.println(frames.size());
		for (int i = 0; i < frames.size(); i++) {
			System.out.println(frames.get(i).getAttribute("name"));
			String frameNametwo = frames.get(i).getAttribute("name");
			driver.switchTo().defaultContent();
			Thread.sleep(1000);
			driver.switchTo().frame(frameNametwo);
			try {
				if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
					System.out.println("Element" + xpath + "is present in the frame: " + frameNametwo);
					frameNameone = frameNametwo;
					break;
				}
			} catch (Exception e) {
				System.out.println("Element is not present inside the frame: " + xpath);
				driver.switchTo().defaultContent();
			}
		}
		System.out.println("In GoToFrames Method End");
		return frameNameone;
	}

	/* Method to switch to a frame using tag name and name locator */
	public static String goToFrameByTag_NameByName(String name) throws InterruptedException {
		System.out.println("In GoToFrames Method Start");
		String frameNameone = null;
		List<WebElement> frames = driver.findElements(By.tagName("iframe"));
		System.out.println(frames.size());
		for (int i = 0; i < frames.size(); i++) {
			System.out.println(frames.get(i).getAttribute("name"));
			String frameNametwo = frames.get(i).getAttribute("name");
			driver.switchTo().defaultContent();
			Thread.sleep(1000);
			driver.switchTo().frame(frameNametwo);
			try {
				if (driver.findElement(By.name(name)).isDisplayed()) {
					System.out.println("Element" + name + "is present in the frame: " + frameNametwo);
					frameNameone = frameNametwo;
					break;
				}
			} catch (Exception e) {
				System.out.println("Element is not present inside the frame: " + name);
				driver.switchTo().defaultContent();
			}
		}
		System.out.println("In GoToFrames Method End");
		return frameNameone;
	}

	/* Method to switch to a frame using tag name and id locator */
	public static String goToFrameByTag_NameById(String id) throws InterruptedException {
		System.out.println("In GoToFrames Method Start");
		String frameNameone = null;
		List<WebElement> frames = driver.findElements(By.tagName("iframe"));
		System.out.println(frames.size());
		for (int i = 0; i < frames.size(); i++) {
			System.out.println(frames.get(i).getAttribute("name"));
			String frameNametwo = frames.get(i).getAttribute("name");
			driver.switchTo().defaultContent();
			Thread.sleep(1000);
			driver.switchTo().frame(frameNametwo);
			try {
				if (driver.findElement(By.id(id)).isDisplayed()) {
					System.out.println("Element" + id + "is present in the frame: " + frameNametwo);
					frameNameone = frameNametwo;
					break;
				}
			} catch (Exception e) {
				System.out.println("Element is not present inside the frame: " + id);
				driver.switchTo().defaultContent();
			}
		}
		System.out.println("In GoToFrames Method End");
		return frameNameone;
	}

	/**********************************************************************************
	 * Generic Methods
	 *****************************************************************/
	/*
	 * Generic method for click on links,radio buttons,check boxes,drop down and
	 * text boxes using xpath locator
	 */
	public static void xpath_GenericMethod_Click(String xpath) throws Exception {
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				WebElement ele = driver.findElement(By.xpath(xpath));
				waitForElement(ele, 1);
				ele.click();
				 waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.xpath(xpath));
				waitForElement(ele1, 1);
				ele1.click();
				 waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.xpath(xpath));
				waitForElement(ele, 1);
				ele.click();
				 waitFor(1);
			}
		}
	}

	/*
	 * Generic method for click on links,radio buttons,check boxes,drop down and
	 * text boxes using id locator
	 */
	public static void id_GenericMethod_Click(String id) throws Exception {
		try {
			if (driver.findElement(By.id(id)).isDisplayed()) {
				WebElement ele = driver.findElement(By.id(id));
				waitForElement(ele, 1);
				ele.click();
				 waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameById(id);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdById(id);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.id(id));
				waitForElement(ele1, 1);
				ele1.click();
				 waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.id(id));
				waitForElement(ele, 1);
				ele.click();
				 waitFor(1);
			}
		}
	}

	/*
	 * Generic method for click on links,radio buttons,check boxes,drop down and
	 * text boxes using name locator
	 */
	public static void name_GenericMethod_Click(String name) throws Exception {
		try {
			if (driver.findElement(By.name(name)).isDisplayed()) {
				WebElement ele = driver.findElement(By.name(name));
				waitForElement(ele, 1);
				ele.click();
				 waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByName(name);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByName(name);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.name(name));
				waitForElement(ele1, 1);
				ele1.click();
				waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.name(name));
				waitForElement(ele, 1);
				ele.click();
				 waitFor(1);
			}
		}
	}

	/* Generic method for clear on text boxes and text area using xpath locator */
	public static void xpath_GenericMethod_Clear(String xpath) throws Exception {
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				WebElement ele = driver.findElement(By.xpath(xpath));
				waitForElement(ele, 1);
				ele.clear();
				 waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.xpath(xpath));
				waitForElement(ele1, 1);
				ele1.clear();
				 waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.xpath(xpath));
				waitForElement(ele, 1);
				ele.clear();
				waitFor(1);
			}
		}
	}

	/* Generic method for clear on text boxes and text area using id locator */
	public static void id_GenericMethod_Clear(String id) throws Exception {
		try {
			if (driver.findElement(By.id(id)).isDisplayed()) {
				WebElement ele = driver.findElement(By.id(id));
				waitForElement(ele, 1);
				ele.clear();
				waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameById(id);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdById(id);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.id(id));
				waitForElement(ele1, 1);
				ele1.clear();
				waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.id(id));
				waitForElement(ele, 1);
				ele.clear();
				 waitFor(1);
			}
		}
	}

	/* Generic method for clear on text boxes and text area using name locator */
	public static void name_GenericMethod_Clear(String name) throws Exception {
		try {
			if (driver.findElement(By.name(name)).isDisplayed()) {
				WebElement ele = driver.findElement(By.name(name));
				waitForElement(ele, 1);
				ele.clear();
				 waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByName(name);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByName(name);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.name(name));
				waitForElement(ele1, 1);
				ele1.clear();
				 waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.name(name));
				waitForElement(ele, 1);
				ele.clear();
				waitFor(1);
			}
		}
	}

	/*
	 * Generic method for sendkeys on text boxes and text area using xpath locator
	 */
	public static void xpath_GenericMethod_Sendkeys(String xpath, String testdata) throws Exception {
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				System.out.println("Element" + xpath + "is present outside the frame: ");
				WebElement ele = driver.findElement(By.xpath(xpath));
				waitForElement(ele, 1);
				ele.sendKeys(testdata);
				 waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.xpath(xpath));
				waitForElement(ele1, 1);
				ele1.sendKeys(testdata);
				waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.xpath(xpath));
				waitForElement(ele, 1);
				ele.sendKeys(testdata);
				 waitFor(1);
			}
		}
	}

	/* Generic method for sendkeys on text boxes and text area using id locator */
	public static void id_GenericMethod_Sendkeys(String id, String testdata) throws Exception {
		try {
			if (driver.findElement(By.id(id)).isDisplayed()) {
				System.out.println("Element" + id + "is present outside the frame: ");
				WebElement ele = driver.findElement(By.id(id));
				waitForElement(ele, 1);
				ele.sendKeys(testdata);
				waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameById(id);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdById(id);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.id(id));
				waitForElement(ele1, 1);
				ele1.sendKeys(testdata);
				waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.id(id));
				waitForElement(ele, 1);
				ele.sendKeys(testdata);
				waitFor(1);
			}
		}
	}

	/* Generic method for sendkeys on text boxes and text area using name locator */
	public static void name_GenericMethod_Sendkeys(String name, String testdata) throws Exception {
		try {
			if (driver.findElement(By.name(name)).isDisplayed()) {
				System.out.println("Element" + name + "is present outside the frame: ");
				WebElement ele = driver.findElement(By.name(name));
				waitForElement(ele, 1);
				ele.sendKeys(testdata);
				 waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByName(name);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByName(name);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.name(name));
				waitForElement(ele1, 1);
				ele1.sendKeys(testdata);
				waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.name(name));
				waitForElement(ele, 1);
				ele.sendKeys(testdata);
				 waitFor(1);
			}
		}
	}

	public static void xpath_GenericMethod_selectFromDropdownUsingValuebyclickingOnDropdown(String expath,
			String value) {
		try {
			if (driver.findElement(By.xpath(expath)).isDisplayed()) {
				Select s = new Select(driver.findElement(By.xpath(expath)));
				s.selectByValue(value);
				waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(expath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(expath);
					System.out.println("using getattribute Id");
				}
				Select s = new Select(driver.findElement(By.xpath(expath)));
				s.selectByValue(value);
				 waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				Select s = new Select(driver.findElement(By.xpath(expath)));
				s.selectByValue(value);
				 waitFor(1);
			}
		}
	}

	public static void xpath_GenericMethod_selectFromDropdownUsingVisibleTextbyclickingOnDropdown(String expath,
			String valuetoselect) {
		try {
			if (driver.findElement(By.xpath(expath)).isDisplayed()) {
				Select s = new Select(driver.findElement(By.xpath(expath)));
				s.selectByVisibleText(valuetoselect);
				waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(expath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(expath);
					System.out.println("using getattribute Id");
				}
				Select s = new Select(driver.findElement(By.xpath(expath)));
				s.selectByVisibleText(valuetoselect);
				waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				Select s = new Select(driver.findElement(By.xpath(expath)));
				s.selectByVisibleText(valuetoselect);
				 waitFor(1);
			}
		}
	}

	public static void xpath_GenericMethod_selectFromDropdownUsingIndexbyclickingOnDropdown(String expath,
			int indexvalue) {
		try {
			if (driver.findElement(By.xpath(expath)).isDisplayed()) {
				Select s = new Select(driver.findElement(By.xpath(expath)));
				s.selectByIndex(indexvalue);
				waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(expath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(expath);
					System.out.println("using getattribute Id");
				}
				Select s = new Select(driver.findElement(By.xpath(expath)));
				s.selectByIndex(indexvalue);
				waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				Select s = new Select(driver.findElement(By.xpath(expath)));
				s.selectByIndex(indexvalue);
				waitFor(1);
			}
		}
	}
	/* Generic Method to verify actual and expected text using xpath */

	public static boolean xpath_Genericmethod_VerifyTextEquals(String xpath, String expectedText) throws Exception {
		boolean flag = false;
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				flag = verifyTextEquals(xpath, expectedText);
				waitFor(1);

			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				flag = verifyTextEquals(xpath, expectedText);
				waitFor(1);
				// System.out.println("using get attribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				flag = verifyTextEquals(xpath, expectedText);
				 waitFor(1);
			}
		}
		return flag;
	}

	public static boolean xpath_Genericmethod_VerifyTextContains(String xpath, String expectedText) throws Exception {
		boolean flag = false;
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				flag = verifyTextContains(xpath, expectedText);
				 waitFor(1);

			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				flag = verifyTextContains(xpath, expectedText);
				 waitFor(1);
				// System.out.println("using get attribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				flag = verifyTextContains(xpath, expectedText);
				 waitFor(1);
			}
		}
		return flag;
	}

	public static boolean xpath_Genericmethod_verifyElementPresent(String xpath) throws Exception {
		boolean flag = false;
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				flag = verifyElementPresent(xpath);
				waitFor(1);

			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				flag = verifyElementPresent(xpath);
				 waitFor(1);
				// System.out.println("using get attribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				flag = verifyElementPresent(xpath);
				waitFor(1);
			}
		}
		return flag;
	}

	public static String xpath_Genericmethod_getFirstSelectedValueFromDropdown(String xpath) throws Exception {
		String textvalue = null;
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				textvalue = getFirstSelectedValueFromDropdown(xpath);
				 waitFor(1);

			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				textvalue = getFirstSelectedValueFromDropdown(xpath);
				 waitFor(1);
				// System.out.println("using get attribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				textvalue = getFirstSelectedValueFromDropdown(xpath);
				 waitFor(1);
			}
		}
		return textvalue;
	}

	public static List<String> xpath_Genericmethod_getAllDropDownValues(String xpath) throws Exception {
		List<String> ddvalues = new LinkedList<String>();
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				ddvalues = getAllDropDownValues(xpath);
				waitFor(1);

			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				ddvalues = getAllDropDownValues(xpath);
				waitFor(1);
				// System.out.println("using get attribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				ddvalues = getAllDropDownValues(xpath);
			 waitFor(1);
			}
		}
		return ddvalues;
	}

	public static String xpath_Genericmethod_readValueFromInput(String xpath) throws Exception {
		String value = null;
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				value = readValueFromInput(xpath);
				waitFor(1);
			}
		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				value = readValueFromInput(xpath);
				waitFor(1);
				// System.out.println("using get attribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				value = readValueFromInput(xpath);
				 waitFor(1);
			}
		}
		return value;
	}

	public static String xpath_Genericmethod_getElementText(String xpath) throws Exception {
		String text = null;
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				text = getElementText(xpath);
				waitFor(1);

			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				text = getElementText(xpath);
				waitFor(1);
				// System.out.println("using get attribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				text = getElementText(xpath);
				 waitFor(1);
			}
		}
		return text;
	}

	public static void xpath_Genericmethod_HandleStaleElement(String xpath) throws Exception {

		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				HandleStaleElement(xpath);
				waitFor(1);

			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				HandleStaleElement(xpath);
				waitFor(1);
				// System.out.println("using get attribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				HandleStaleElement(xpath);
				 waitFor(1);
			}
		}

	}

	public static void xpath_GenericMethod_ClickWBResultsRow(String caseid, String xpathstart, String xpathend, int j)
			throws Exception {
		try {
			if (driver.findElement(By.xpath(xpathstart + j + xpathend)).isDisplayed()) {
				Generic_clickWBResults_Row(caseid, xpathstart, xpathend, j);

				waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpathstart + j + xpathend);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpathstart + j + xpathend);
					System.out.println("using getattribute Id");
				}
				Generic_clickWBResults_Row(caseid, xpathstart, xpathend, j);
				 waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("sample test");
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				Generic_clickWBResults_Row(caseid, xpathstart, xpathend, j);
				 waitFor(1);
			}
		}
	}
	
	
	public static void xpath_GenericMethod_HoverOnDemoScreenPops(String xpath) throws Exception {
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				WebElement ele=driver.findElement(By.xpath(xpath));
				waitForElement(ele, 1);
				Actions action=new Actions(driver);
				action.moveToElement(ele).build().perform();
				 waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.xpath(xpath));
				waitForElement(ele1, 1);
				Actions action=new Actions(driver);
				action.moveToElement(ele1).build().perform();
				 waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.xpath(xpath));
				waitForElement(ele, 1);
				Actions action=new Actions(driver);
				action.moveToElement(ele).build().perform();
				 waitFor(1);
			}
		}
	}
	public static void xpath_GenericMethod_scrollIntoView(String xpath) throws Exception {
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				WebElement ele = driver.findElement(By.xpath(xpath));
				executeScript("arguments[0].scrollIntoView();", ele);
				 waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.xpath(xpath));
				waitForElement(ele1, 1);
				executeScript("arguments[0].scrollIntoView();", ele1);
				 waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.xpath(xpath));
				executeScript("arguments[0].scrollIntoView();", ele);
				 waitFor(1);
			}
		}
	}
	
	public static void xpath_GenericMethod_ActionClass(String xpath) throws Exception {
		try {
			if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
				WebElement ele = driver.findElement(By.xpath(xpath));
				
				  Actions move = new Actions(driver); move.moveToElement(ele).clickAndHold();
				  move.moveByOffset(125,0); move.release(); move.perform();
				 
			
				 waitFor(1);
			}

		} catch (Exception e) {
			try {
				String framebytagName = goToFrameByTag_NameByXpath(xpath);

				System.out.println("using getattribute name");
				if (framebytagName == null) {
					goToFrameByTag_IdByXpath(xpath);
					System.out.println("using getattribute Id");
				}
				WebElement ele1 = driver.findElement(By.xpath(xpath));
				waitForElement(ele1, 1);
				
				  Actions move = new Actions(driver); move.moveToElement(ele1).clickAndHold();
				  move.moveByOffset(125,0); move.release(); move.perform();
				 
				
				 waitFor(1);
				// System.out.println("using getattribute Id");
				// driver.switchTo().defaultContent();
			} catch (Exception e1) {
				System.out.println("switching to default content");
				driver.switchTo().defaultContent();
				WebElement ele = driver.findElement(By.xpath(xpath));
				
				  Actions move = new Actions(driver); move.moveToElement(ele).clickAndHold();
				  move.moveByOffset(125,0); move.release(); move.perform();
				 
			
				 waitFor(1);
			}
		}
	}
	
	public static void xpath_GenericMethod_getText_TableResultsRow(String expectedtext, String xpathstart, String xpathend, int j)
            throws Exception {
     try {
            if (driver.findElement(By.xpath(xpathstart + j + xpathend)).isDisplayed()) {
                  Generic_getText_TableResults_Row(expectedtext, xpathstart, xpathend, j);

                  waitFor(1);
            }

     } catch (Exception e) {
            try {
                  String framebytagName = goToFrameByTag_NameByXpath(xpathstart + j + xpathend);

                  System.out.println("using getattribute name");
                  if (framebytagName == null) {
                         goToFrameByTag_IdByXpath(xpathstart + j + xpathend);
                         System.out.println("using getattribute Id");
                  }
                  Generic_getText_TableResults_Row(expectedtext, xpathstart, xpathend, j);
                  waitFor(1);
                  // System.out.println("using getattribute Id");
                  // driver.switchTo().defaultContent();
            } catch (Exception e1) {
                  System.out.println("sample test");
                  System.out.println("switching to default content");
                  driver.switchTo().defaultContent();
                  Generic_getText_TableResults_Row(expectedtext, xpathstart, xpathend, j);
                  waitFor(1);
            }
     }
}


	
	
	/*****************************************************************************************
	 * Generic Methods End
	 ***************************************************************************************************/
	/*
	 * Method to wait for an element until it is visible for the specified amount of
	 * time
	 */
	
	public static void Generic_getText_TableResults_Row(String expectedtext, String xpathstart, String xpathend, int j) {
        for (int i = j; i <= 5; i++) {
               WebElement actualtext = driver.findElement(By.xpath(xpathstart + i + xpathend));
              // try {
            	   Reporter.log("Before Assertion: Actual Text is: "+actualtext.getText()+" Expected Text is: "+expectedtext);
            	   Assert.assertTrue(actualtext.getText().equalsIgnoreCase(expectedtext),"Actual and Expected are not same");
            	   Reporter.log("After Assertion: Actual Text is: "+actualtext.getText()+" Expected Text is: "+expectedtext);
				/*
				 * if (actualtext.getText().equalsIgnoreCase(expectedtext)) {
				 * scrollToElement(actualtext);
				 * Reporter.log("Actual Text is: "+actualtext.getText()+" Expected Text is: "
				 * +expectedtext);
				 * 
				 * // break; }else {
				 * Reporter.log("Actual Text is: "+actualtext.getText()+" Expected Text is: "
				 * +expectedtext); }
				 */
            /*   } catch (Exception e) {
            	   Reporter.log("Expected text not Found: "+expectedtext);
                     System.out.println("No such element" + e.getMessage());
               }*/
        }
        // driver.switchTo().defaultContent();
 }

	

	public static void waitForElement(WebElement element, int timeunitForsec) {
		WebDriverWait wait = new WebDriverWait(driver, timeunitForsec);
		wait.until(ExpectedConditions.visibilityOf(element));
	}

	/* Method to wait for specified amount of time */
	public static void waitFor(int millisec) {
		try {
			Thread.sleep(millisec * 1000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

	/*
	 * Method to select value from drop down using visible text and xpath locator of
	 * WebElement
	 */
	public static void xpath_SelectFromDropdownUsingVisibleText(String expath, String valuetoselect) {
		Select s = new Select(driver.findElement(By.xpath(expath)));
		s.selectByVisibleText(valuetoselect);
	}

	/*
	 * Method to select value from drop down using visible text and xpath locator of
	 * WebElement
	 */
	public static void name_SelectFromDropdownUsingVisibleText(String name, String valuetoselect) {
		Select s = new Select(driver.findElement(By.name(name)));
		s.selectByVisibleText(valuetoselect);
	}

	/*
	 * Method to select value from drop down using visible text and id locator of
	 * WebElement
	 */
	public static void id_SelectFromDropdownUsingVisibleText(String id, String valuetoselect) {
		Select s = new Select(driver.findElement(By.id(id)));
		s.selectByVisibleText(valuetoselect);
	}

	/*
	 * Method to select value from drop down using index value and xpath locator of
	 * WebElement
	 */
	public static void xpath_SelectFromDropdownUsingIndex(String expath, int indexvalue) {
		Select s = new Select(driver.findElement(By.xpath(expath)));
		s.selectByIndex(indexvalue);
	}

	/*
	 * Method to select value from drop down using index value and name locator of
	 * WebElement
	 */
	public static void name_SelectFromDropdownUsingIndex(String name, int indexvalue) {
		Select s = new Select(driver.findElement(By.name(name)));
		s.selectByIndex(indexvalue);
	}

	/*
	 * Method to select value from drop down using index value and id locator of
	 * WebElement
	 */
	public static void id_SelectFromDropdownUsingIndex(String id, int indexvalue) {
		Select s = new Select(driver.findElement(By.id(id)));
		s.selectByIndex(indexvalue);
	}

	/*
	 * Method to select value from drop down using value and xpath locator of
	 * WebElement
	 */
	public static void xpath_SelectFromDropdownUsingValue(String expath, String value) {
		Select s = new Select(driver.findElement(By.xpath(expath)));
		s.selectByValue(value);
	}

	/*
	 * Method to select value from drop down using value and name locator of
	 * WebElement
	 */
	public static void name_SelectFromDropdownUsingValue(String name, String value) {
		Select s = new Select(driver.findElement(By.name(name)));
		s.selectByValue(value);
	}

	/*
	 * Method to select value from drop down using value and id locator of
	 * WebElement
	 */
	public static void id_SelectFromDropdownUsingValue(String id, String value) {
		Select s = new Select(driver.findElement(By.id(id)));
		s.selectByValue(value);
	}

	/**********************************************************************************
	 * New Methods Added
	 *****************************************************************/
	public static synchronized boolean verifyElementPresent(String xpath) {
		boolean isDisplayed = false;
		try {
			isDisplayed = driver.findElement(By.xpath(xpath)).isDisplayed();
			System.out.println(driver.findElement(By.xpath(xpath)).getText() + " Element is displayed");
		} catch (Exception ex) {
			System.out.println("Element not found " + ex);
		}

		return isDisplayed;
	}

	public static synchronized boolean verifyElementNotPresent(WebElement element) {
		boolean isDispalyed = false;
		try {
			element.isDisplayed();
			System.out.println(element.getText() + " is dispalyed");
		} catch (Exception ex) {
			System.out.println("Element not found " + ex);
			isDispalyed = true;
		}

		return isDispalyed;
	}

	public static synchronized boolean verifyTextEquals(String xpath, String expectedText) {
		boolean flag = false;
		try {
			String actualText = driver.findElement(By.xpath(xpath)).getText();
			if (actualText.equals(expectedText)) {
				System.out.println("actualText is : " + actualText + " and expected text is: " + expectedText);
				Reporter.log("Actual text is: "+actualText+" "+"and Expected text is: "+expectedText);
				return flag = true;
			} else {
				System.out.println("actualText is : " + actualText + " and  expected text is: " + expectedText);
				Reporter.log("Actual text is: "+actualText+" "+"and Expected text is: "+expectedText);
				return flag;
			}
		} catch (Exception ex) {
			System.out.println("actualText is :" + driver.findElement(By.xpath(xpath)).getText()
					+ " and expected text is: " + expectedText);
			System.out.println("text not matching" + ex);
			Reporter.log("actualText is :" + driver.findElement(By.xpath(xpath)).getText()
					+ " and expected text is: " + expectedText);
			return flag;
		}
	}

	public static synchronized boolean verifyTextContains(String xpath, String expectedText) {
		boolean flag = false;
		try {
			String actualText = driver.findElement(By.xpath(xpath)).getText();
			if (actualText.contains(expectedText)) {
				System.out.println("actualText is : " + actualText + " and expected text is: " + expectedText);
				return flag = true;
			} else {
				System.out.println("actualText is : " + actualText + " and  expected text is: " + expectedText);
				return flag;
			}
		} catch (Exception ex) {
			System.out.println("actualText is :" + driver.findElement(By.xpath(xpath)).getText()
					+ " and expected text is: " + expectedText);
			System.out.println("text not matching" + ex);
			return flag;
		}
	}

	public static Object executeScript(String script, Object... args) {
		JavascriptExecutor exe = (JavascriptExecutor) driver;
		System.out.println(script);
		return exe.executeScript(script, args);
	}

	public static Object executeScript(String script) {
		JavascriptExecutor exe = (JavascriptExecutor) driver;
		System.out.println(script);
		return exe.executeScript(script);
	}

	public static void scrollToElement(WebElement element) {
		executeScript("window.scrollTo(arguments[0],arguments[1])", element.getLocation().x, element.getLocation().y);
		System.out.println(element);
	}

	public static void scrollToElemetAndClick(WebElement element) {
		scrollToElement(element);
		element.click();
		System.out.println(element);
	}

	public static void scrollIntoView(WebElement element) {
		executeScript("arguments[0].scrollIntoView();", element);
		System.out.println(element);
	}
	
	

	public static void scrollIntoViewAndClick(WebElement element) {
		scrollIntoView(element);
		element.click();
		System.out.println(element);
	}

	public static void scrollDownVertically() {
		executeScript("window.scrollTo(0, document.body.scrollHeight)");
	}

	public static void scrollUpVertically() {
		executeScript("window.scrollTo(0, -document.body.scrollHeight)");
	}

	public static void scrollDownByPixel() {
		executeScript("window.scrollBy(0,1500)");
	}

	public static void scrollUpByPixel() {
		executeScript("window.scrollBy(0,-1500)");
	}

	public static void ZoomInBypercentage() {
		executeScript("document.body.style.zoom='40%'");
	}

	public static void ZoomBy100percentage() {
		executeScript("document.body.style.zoom='100%'");
	}

	public static String getFirstSelectedValueFromDropdown(String xpath) {
		String value = new Select(driver.findElement(By.xpath(xpath))).getFirstSelectedOption().getText();
		System.out.println("WebELement : " + driver.findElement(By.xpath(xpath)) + " Value : " + value);
		return value;
	}

	public static List<String> getAllDropDownValues(String xpath) {
		Select select = new Select(driver.findElement(By.xpath(xpath)));
		List<WebElement> elementList = select.getOptions();
		List<String> valueList = new LinkedList<String>();

		for (WebElement element : elementList) {
			System.out.println(element.getText());
			valueList.add(element.getText());
		}
		return valueList;
	}

	public String readValueFromElement(WebElement element) {

		if (null == element) {
			log.info("weblement is null");
			return null;
		}

		boolean displayed = false;
		try {
			displayed = isDisplayed(element);
		} catch (Exception e) {
			log.error(e);
			return null;
		}

		if (!displayed)
			return null;
		String text = element.getText();
		log.info("weblement valus is.." + text);
		return text;
	}

	public static String readValueFromInput(String xpath) {
		if (null == driver.findElement(By.xpath(xpath)))
			return null;
		if (!isDisplayed(driver.findElement(By.xpath(xpath))))
			return null;
		String value = driver.findElement(By.xpath(xpath)).getAttribute("value");
		return value;
	}

	public static boolean isDisplayed(WebElement element) {
		try {
			element.isDisplayed();
			System.out.println("element is displayed.." + element);
			return true;
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return false;
		}
	}

	public static boolean isNotDisplayed(WebElement element) {
		try {
			element.isDisplayed();
			System.out.println("element is displayed.." + element);
			return false;
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return true;
		}
	}

	public static String getDisplayText(WebElement element) {
		if (null == element)
			return null;
		if (!isDisplayed(element))
			return null;
		return element.getText();
	}

	public static synchronized String getElementText(WebElement element) {
		if (null == element) {
			System.out.println("weblement is null");
			return null;
		}
		String elementText = null;
		try {
			elementText = element.getText();
		} catch (Exception ex) {
			System.out.println("Element not found " + ex);
		}
		return elementText;
	}

	public static synchronized String getElementText(String xpath) {
		if (null == driver.findElement(By.xpath(xpath))) {
			System.out.println("webelement is null");
			return null;
		}
		String elementText = null;
		try {
			elementText = driver.findElement(By.xpath(xpath)).getText();
		} catch (Exception ex) {
			System.out.println("Element not found " + ex);
		}
		return elementText;
	}

	public static void Generic_clickWBResults_Row(String caseid, String xpathstart, String xpathend, int j) {
		for (int i = j; i <= 100; i++) {
			WebElement caseID = driver.findElement(By.xpath(xpathstart + i + xpathend));
			try {
				if (caseID.getText().equalsIgnoreCase(caseid)) {
					
					caseID.click();
					break;
				}
			} catch (Exception e) {
				System.out.println("No such element" + e.getMessage());
			}
		}
		// driver.switchTo().defaultContent();
	}
}
