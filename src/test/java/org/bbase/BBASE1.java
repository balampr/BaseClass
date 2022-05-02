package org.bbase;

import java.awt.AWTException;
import java.awt.Desktop.Action;
import java.awt.Robot;
import java.awt.event.ActionEvent;
import java.awt.event.KeyEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.xml.transform.Source;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BBASE1 {

	public static WebDriver driver;
	public static WebDriver Actions;

	public static void openChromeBrowser() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();

	}

	public static void launchUrl(String url) {
		driver.get(url);

	}

	public static void maxWindow() {
		driver.manage().window().maximize();

	}

	public static void toHold(int time) throws InterruptedException {
		Thread.sleep(time);
	}

	public static void toFillTextBox(WebElement element, String pass) {
		element.sendKeys(pass);

	}

	public static void fetchTitle() {
		String title = driver.getTitle();
		System.out.println(title);

	}

	public static void fetchCurrentUrl() {
		String currentUrl = driver.getCurrentUrl();
		System.out.println(currentUrl);

	}

	public static void click(WebElement name) {
		name.click();
	}

	public static void closeBrowser() {
		driver.quit();

	}

	public static void moveToElement(WebElement target) {
		Actions ac = new Actions(driver);
		ac.moveToElement(target).perform();

	}

	public static void dragAndDrop(WebElement src, WebElement dest) {
		Actions a = new Actions(driver);
		a.dragAndDrop(src, dest).perform();
	}

	public static void rightClick(WebElement element) {
		Actions ac = new Actions(driver);
		ac.contextClick(element).perform();
	}

	public static void doubleClick(WebElement element) {
		Actions ac = new Actions(driver);
		ac.doubleClick(element).perform();
	}

	public static void copy() throws AWTException {
		Robot rc = new Robot();
		rc.keyPress(KeyEvent.VK_CONTROL);
		rc.keyPress(KeyEvent.VK_C);
		rc.keyRelease(KeyEvent.VK_C);
		rc.keyRelease(KeyEvent.VK_CONTROL);
	}

	public static void paste() throws AWTException {
		Robot rc = new Robot();
		rc.keyPress(KeyEvent.VK_CONTROL);
		rc.keyPress(KeyEvent.VK_V);
		rc.keyRelease(KeyEvent.VK_V);
		rc.keyRelease(KeyEvent.VK_CONTROL);
	}

	public static void takeScreenShot(String picname) throws IOException {
		TakesScreenshot tk = (TakesScreenshot) driver;
		File src = tk.getScreenshotAs(OutputType.FILE);
		File dest = new File("F:\\Maven\\BASE12\\" + picname + ".png");
		FileUtils.copyFile(src, dest);
	}

	public static void scrlDw(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].scrollIntoView(true)", element);

	}

	public static void scrlUp(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].scrollIntoView(false)", element);

	} 
	public static void readExcelSheet(String path,String name) throws IOException {
		File f = new File(path);
		FileInputStream fi = new FileInputStream(f);
		Workbook book = new XSSFWorkbook(fi);
		Sheet sh = book.getSheet(name);
		for (int i = 0; i < sh.getPhysicalNumberOfRows(); i++) {
			Row row = sh.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				int Type = cell.getCellType();
				if(Type==1) {
					String s = cell.getStringCellValue();
					System.out.println(s);
					}
				else if(DateUtil.isCellDateFormatted(cell)){
					Date da = cell.getDateCellValue();
						SimpleDateFormat sim = new SimpleDateFormat("dd/MMMMM/yy");
						System.out.println(sim);
				}
				else {
					double n = cell.getNumericCellValue();
					long l =(long)n;
					    String v = String.valueOf(l);
					    System.out.println(v);
					    
				}
				
				}
			
			}
			
			
			
			
		}
		

	}


