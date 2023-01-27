package org.baseclass;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.Set;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {

	public static WebDriver driver;
	public static Actions act;
	public static Robot robo;
	public static Alert alt;
	public static JavascriptExecutor js;
	public static TakesScreenshot screenShot;
	public static Select sel;

	public static void openChromeBrowser() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
	}

	public static void openFireFoxBrowser() {
		WebDriverManager.firefoxdriver().avoidBrowserDetection().setup();
		driver = new FirefoxDriver();
	}

	public static void openEdgeBrowser() {
		WebDriverManager.edgedriver().setup();
		driver = new EdgeDriver();
	}

	public static void threadSleep(int time) throws InterruptedException {
		Thread.sleep(time);
	}

	public static void launchUrl(String url) {
		driver.get(url);
	}

	public static String pageTitle() {
		return driver.getTitle();

	}

	public static void pageUrl() {
		String pageUrl = driver.getCurrentUrl();
		System.out.println(pageUrl);

	}

	public static void closeTab() {
		driver.close();

	}

	public static void quitBrowser() {
		driver.quit();
	}

	public static void maximizeWindow() {
		driver.manage().window().maximize();
	}

	public static void clearTextBox(WebElement elementRef) {
		elementRef.clear();
	}
	
	
	public static void implictWait(int seconds) {
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(seconds));

	}

	public static void passValue(WebElement elementRef, String value) {
		elementRef.sendKeys(value);
	}

	public static void clickButton(WebElement elementRef) {
		elementRef.click();
	}

	public static String fetchBackUserInput(WebElement elementRef) {
		return elementRef.getAttribute("value");

	}

	public static String fetchBackAttributeValue(WebElement elementRef, String AttributeName) {
		return elementRef.getAttribute(AttributeName);

	}

	public static String fetchBackText(WebElement elementRef) {
		return elementRef.getText();

	}

	public static WebElement mouseHovering(WebElement elementRef) {
		act = new Actions(driver);
		act.moveToElement(elementRef).perform();
		return elementRef;
	}

	public static void dragAndDrop(WebElement src, WebElement trgt) {
		act = new Actions(driver);
		act.dragAndDrop(src, trgt).perform();

	}

	public static void doubleclick(WebElement elementRef) {
		act = new Actions(driver);
		act.doubleClick(elementRef).perform();

	}

	public static void rightClick(WebElement elementRef) {
		act = new Actions(driver);
		act.contextClick(elementRef).perform();

	}

	public static void keyDownUp(Keys keys) {
		act = new Actions(driver);
		act.keyDown(keys).perform();
		act.keyUp(keys).perform();
	}

	public static void keyPressRelease(int keyEvent) throws AWTException {
		robo = new Robot();
		robo.keyPress(keyEvent);
		robo.keyRelease(keyEvent);

	}

	public static void copy() throws AWTException {
		robo = new Robot();
		robo.keyPress(KeyEvent.VK_CONTROL);
		robo.keyPress(KeyEvent.VK_C);
		robo.keyRelease(KeyEvent.VK_CONTROL);
		robo.keyRelease(KeyEvent.VK_C);

	}

	public static void paste() throws AWTException {
		robo = new Robot();
		robo.keyPress(KeyEvent.VK_CONTROL);
		robo.keyPress(KeyEvent.VK_V);
		robo.keyRelease(KeyEvent.VK_CONTROL);
		robo.keyRelease(KeyEvent.VK_V);

	}

	public static void selectAll() throws AWTException {
		robo = new Robot();
		robo.keyPress(KeyEvent.VK_CONTROL);
		robo.keyPress(KeyEvent.VK_A);
		robo.keyRelease(KeyEvent.VK_CONTROL);
		robo.keyRelease(KeyEvent.VK_A);
	}

	public static void acceptAlert() {
		driver.switchTo().alert();
		alt.accept();
	}

	public static void dismissAlert() {
		driver.switchTo().alert();
		alt.dismiss();

	}

	public static void passValueToAlert(String input) {
		driver.switchTo().alert();
		alt.sendKeys(input);

	}

	public static void fetchBackAlertText() {
		driver.switchTo().alert();
		String alertTxt = alt.getText();
		System.out.println(alertTxt);

	}

	public static void jsPassValue(String input, WebElement elementRef) {
		js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].setAttribute('value'," + "'" + input + "')", elementRef);

	}

	public static void jsFetchBackInput() {
		js = (JavascriptExecutor) driver;
		Object input = js.executeScript("return arguments[0].getAttribute('value')");
		String text = (String) input;
		System.out.println(text);
	}

	public static void jsFetchBackAttributeValue(String attributeName) {
		js = (JavascriptExecutor) driver;
		js.executeScript("return arguments[0].setAttribute('value'," + "'" + attributeName + "')");
	}

	public static void jsClick(WebElement elementRef) {
		js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].click()", elementRef);
	}

	public static void jsScrollDown(WebElement elementRef) {
		js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].scrollIntoView(true)", elementRef);
	}

	public static void jsScrollUp(WebElement elementRef) {
		js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].scrollIntoView(false)", elementRef);

	}

	public static void switchToFrame(String id) {
		driver.switchTo().frame(id);

	}

	public static void switchToFrameWebElement(WebElement elementRef) {
		driver.switchTo().frame(elementRef);

	}

	public static void switchToFrameIndex(int index) {
		driver.switchTo().frame(index);

	}

	public static void switchToParentFrame() {
		driver.switchTo().parentFrame();

	}

	public static void switchToMainHtml() {
		driver.switchTo().defaultContent();

	}

	public static void takeScreenShot(String fileName) throws IOException {
		screenShot = (TakesScreenshot) driver;
		File temp = screenShot.getScreenshotAs(OutputType.FILE);
		File system = new File("D:\\Eclipse\\MavenProject\\Screenshots\\" + fileName + ".PNG");
		FileUtils.copyFile(temp, system);
		System.out.println(fileName + " ScreenShot Taken");

	}

	public static boolean checkDropDown(WebElement elementRef) {
		sel = new Select(elementRef);
		boolean multiple = sel.isMultiple();
		return multiple;
	}

	public static void selectByValue(WebElement elementRef, String value) {
		sel = new Select(elementRef);
		sel.selectByValue(value);
	}

	public static void selectByText(WebElement elementRef, String text) {
		sel = new Select(elementRef);
		sel.selectByVisibleText(text);

	}

	public static void selectByIndex(WebElement elementRef, int index) {
		sel = new Select(elementRef);
		sel.selectByIndex(index);

	}

	public static List<WebElement> getAllOptions(WebElement elementRef) {
		sel = new Select(elementRef);
		return sel.getOptions();

	}

	public static String getSelectedOptionInSingleDropDown(WebElement elementRef) {
		sel = new Select(elementRef);

		List<WebElement> selectedOptions = sel.getAllSelectedOptions();
		return selectedOptions.get(0).getText();

	}

	public static void getFirstSelectedOption(WebElement elementRef) {
		sel = new Select(elementRef);

		WebElement firstSelected = sel.getFirstSelectedOption();
		String text = firstSelected.getText();
		System.out.println("First Selected Option : " + text);

	}

	public static void deelectByValue(WebElement elementRef, String value) {
		sel = new Select(elementRef);
		sel.deselectByValue(value);
	}

	public static void deselectByText(WebElement elementRef, String text) {
		sel = new Select(elementRef);
		sel.deselectByVisibleText(text);

	}

	public static void deselectByIndex(WebElement elementRef, int index) {
		sel = new Select(elementRef);
		sel.deselectByIndex(index);

	}

	public static void deselectAll(WebElement elementRef) {
		sel = new Select(elementRef);
		sel.deselectAll();

	}

	public static String getParentWindowId() {
		String parentId = driver.getWindowHandle();
		return parentId;
	}

	public static Set<String> getAllWindowId() {
		Set<String> allWindows = driver.getWindowHandles();
		return allWindows;

	}

	public static void switchToWindow(String windowId) {
		driver.switchTo().window(windowId);
	}

	public static void switchToWindow(Set<String> refName, int index) {
		List<String> list = new ArrayList<>(refName);
		String singleWind = list.get(index);
		driver.switchTo().window(singleWind);

	}

	public static String readDataFromExcel(String fileName, String sheetName, int row, int cellNo) throws IOException {

		File excel = new File("D:\\Eclipse\\MavenProject\\ExcelSheet\\" + fileName + ".xlsx");

		FileInputStream read = new FileInputStream(excel);

		Workbook book = new XSSFWorkbook(read);

		Sheet sheet = book.getSheet(sheetName);

		Cell cell = sheet.getRow(row).getCell(cellNo);

		int cellType = cell.getCellType();

		String value;

		if (cellType == 1) {

			 value = cell.getStringCellValue();

		}

		else if (DateUtil.isCellDateFormatted(cell)) {
			Date dateCellValue = cell.getDateCellValue();
			SimpleDateFormat format = new SimpleDateFormat("dd/MM/yy");
			 value = format.format(dateCellValue);

		}

		else {
			double number = cell.getNumericCellValue();

			 value = String.valueOf((long) number);

		}
		return value;

	}

	public static String readDataFromExcel(String fileName, int sheetNo, int row, int cellNo) throws IOException {

		File excel = new File("D:\\Eclipse\\MavenProject\\ExcelSheet\\" + fileName + ".xlsx");

		FileInputStream read = new FileInputStream(excel);

		Workbook book = new XSSFWorkbook(read);
		Sheet sheet = book.getSheetAt(sheetNo);
		Cell cell = sheet.getRow(row).getCell(cellNo);
		int cellType = cell.getCellType();

		String value;

		if (cellType == 1) {

			return value = cell.getStringCellValue();

		}

		else if (DateUtil.isCellDateFormatted(cell)) {
			Date dateCellValue = cell.getDateCellValue();
			SimpleDateFormat format = new SimpleDateFormat("dd/MM/yy");
			 value = format.format(dateCellValue);

		}

		else {
			double number = cell.getNumericCellValue();
			 value = String.valueOf((long) number);

		}
		return value;

	}

	public static void updateDataInExcel(String fileName, String sheetName, int updateRowNo, int updateCellNo,
			String updateValue) throws IOException {

		File excel = new File("D:\\Eclipse\\MavenProject\\ExcelSheet\\" + fileName + ".xlsx");
		FileInputStream read = new FileInputStream(excel);
		Workbook book = new XSSFWorkbook(read);
		Sheet sheet = book.getSheet(sheetName);
		Row row = sheet.getRow(updateRowNo);
		Cell updateCell = row.getCell(updateCellNo);
		updateCell.setCellValue(updateValue);
		FileOutputStream fout = new FileOutputStream(excel);
		book.write(fout);

	}

	public static void createNewExcel(String fileName, String sheetName, int createRowCount) throws IOException {
		File excel = new File("D:\\Eclipse\\MavenProject\\ExcelSheet\\" + fileName + ".xlsx");
		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet(sheetName);
		for (int i = 0; i < createRowCount; i++) {
			sheet.createRow(i);
		}

		FileOutputStream fout = new FileOutputStream(excel);
		book.write(fout);
	}

	public static void createNewSheetInExcel(String fileName, String sheetName, int createRowCount) throws IOException {

		File excel = new File("D:\\Eclipse\\MavenProject\\ExcelSheet\\" + fileName + ".xlsx");
		FileInputStream read = new FileInputStream(excel);
		Workbook book = new XSSFWorkbook(read);
		Sheet sheet = book.createSheet(sheetName);
		for (int i = 0; i < createRowCount; i++) {
			sheet.createRow(i);
		}

		FileOutputStream fout = new FileOutputStream(excel);
		book.write(fout);

	}

	public static void setDataInExcel(String fileName, String sheetName, int RowNo, int CellNo, String Value)
			throws IOException {
		File excel = new File("D:\\Eclipse\\MavenProject\\ExcelSheet\\" + fileName + ".xlsx");
		FileInputStream read = new FileInputStream(excel);
		Workbook book = new XSSFWorkbook(read);
		Sheet sheet1 = book.getSheet(sheetName);
		sheet1.getRow(RowNo).createCell(CellNo).setCellValue(Value);
		FileOutputStream fout = new FileOutputStream(excel);
		book.write(fout);
	}

	public static Date dateAndTime() {
		Date date = new Date();
		return date;
	}

}
