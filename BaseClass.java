package org.utilities;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.reporters.jq.Main;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {

	public static WebDriver driver;

	public static void launchchrome() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
	}

	public static void loadurl(String url) {
		driver.get(url);
	}

	public static void winmax() {
		driver.manage().window().maximize();
	}

	public static void printtitle() {
		String title = driver.getTitle();
		System.out.println(title);
	}

	public static void printcurrenturl() {
		System.out.println(driver.getCurrentUrl());
	}

	public static void fill(WebElement ele, String value) {
		ele.sendKeys(value);

	}

	public static void btnclick(WebElement ele) {
		ele.click();

	}

	public static void printtext(WebElement text) {

		System.out.println(text.getText());

	}

	public static void timeout() throws InterruptedException {
		Thread.sleep(3000);
	}

	public static void implicitwait() {
		driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
	}
	
	public static void mousego(WebElement c) {
		Actions a = new Actions(driver);
		a.moveToElement(c).perform();

	}

	public static void mousemove(WebElement b) {
		Actions a = new Actions(driver);
		a.moveToElement(b).click().perform();
	}

	public static void dragclick(WebElement src,WebElement des) {
		Actions a = new Actions(driver);
		a.dragAndDrop(src, des).click().perform();

	}

	public static void dragmove(WebElement src, WebElement des) {
		Actions a = new Actions(driver);
		a.dragAndDrop(src, des).perform();
	}

	public static void enter(WebElement b) {
		Actions a = new Actions(driver);
		a.click();
	}

	public static void rightclick(WebElement b) {
		Actions a = new Actions(driver);
		a.contextClick().perform();
	}
	public static void keydown(int a) throws AWTException {
		Robot r = new Robot();
		for (int i = 0; i < a; i++) {
			r.keyPress(KeyEvent.VK_DOWN);
			r.keyRelease(KeyEvent.VK_DOWN);
		}
	}

	public static void keyup(int a) throws AWTException {
		Robot r = new Robot();
		for (int i = 0; i < a; i++) {
			r.keyPress(KeyEvent.VK_UP);
			r.keyRelease(KeyEvent.VK_UP);
		}
	}
	public static void keyenter() throws AWTException {
		Robot r = new Robot();
		r.keyPress(KeyEvent.VK_ENTER);
		r.keyRelease(KeyEvent.VK_ENTER);
	}
	public static void keytab() throws AWTException {
		Robot r = new Robot();
		r.keyPress(KeyEvent.VK_TAB);
		r.keyRelease(KeyEvent.VK_TAB);
	}
	
	public static void ctrv() throws AWTException {
		Robot r = new Robot();
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_V);
		r.keyRelease(KeyEvent.VK_CONTROL);
		r.keyRelease(KeyEvent.VK_V);
	}
	public static void ctrx() throws AWTException {
		Robot r = new Robot();
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_X);
		r.keyRelease(KeyEvent.VK_CONTROL);
		r.keyRelease(KeyEvent.VK_X);
	}
	public static void ctra() throws AWTException {
		Robot r = new Robot();
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_A);
		r.keyRelease(KeyEvent.VK_CONTROL);
		r.keyRelease(KeyEvent.VK_A);
		
	}
	
	public static void ctrd() throws AWTException {
		Robot r = new Robot();
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_D);
		r.keyRelease(KeyEvent.VK_CONTROL);
		r.keyRelease(KeyEvent.VK_D);

	}
	public static void simplealert() {
		Alert a = driver.switchTo().alert();
		System.out.println(a.getText());
		a.accept();
	}
	public static void confirmalert(WebElement clk) {
		Alert a = driver.switchTo().alert();
		System.out.println(a.getText());
		a.dismiss();
	}
	public static void promptalert(String s) {
		Alert a = driver.switchTo().alert();
		a.sendKeys(s);
		System.out.println(a.getText());
		a.accept();
	}
	public static void screenshots(String name) throws IOException {
		TakesScreenshot tk =(TakesScreenshot) driver;
		File src = tk.getScreenshotAs(OutputType.FILE);
		File des = new File("C:\\Users\\ARUNKUMAR\\eclipse-workspace\\Maven12PM\\ScreenShots\\" +name+".jpg");
		FileUtils.copyFile(src, des);
	}
	public static void scrollup(WebElement ref) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arugument[0].scrollIntoView(true)",ref );
	}
	public static void scrolldown(WebElement ref) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arugument[0].scrollIntoView(false)",ref);
	}
	public static void jssentskey(WebElement ele,String name) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		String a ="arguments[0].setAttribute('value','sent')";
		String b = a.replace("sent", name);
		js.executeScript(b,ele);
	}
	public static void parentwindow() {
		String a = driver.getWindowHandle();
		System.out.println(a);
	}

	public static void childwindow(int  a) {
		Set<String> allwindows = driver.getWindowHandles();
		int count=0;
		for (String eachid : allwindows) {
			if (count==a) {
				driver.switchTo().window(eachid);
			}
			count++;
		}

	}

	public static void firstwindow(String id) {
		Set<String> allwindow = driver.getWindowHandles();
		for (String eachid : allwindow) {
			if (!id.equals(eachid)) {
				driver.switchTo().window(eachid);
			}
		}
	}

	public static String getdata(int rowNo,int cellNo) throws IOException {
		File f = new File("C:\\Users\\ARUNKUMAR\\eclipse-workspace\\Maven12PM\\ExcelSheet\\ADACTIN.xlsx");

		FileInputStream fin = new FileInputStream(f);

		Workbook w = new XSSFWorkbook(fin);

		Sheet s = w.getSheet("adactin");

		Row row = s.getRow(rowNo);

		Cell cell = row.getCell(cellNo);

		int cellType = cell.getCellType();
		String value="";
		if (cellType==1) {
			value = cell.getStringCellValue();
		}
		else if (cellType==0) {
			if (DateUtil.isCellDateFormatted(cell)) {
				Date d = cell.getDateCellValue();

				SimpleDateFormat sim = new SimpleDateFormat("dd/MM/yyyy");
				value= sim.format(d);
			}
			else {
				double d = cell.getNumericCellValue();

				long l= (long)d;
				value = String.valueOf(l);
			}

		}

		return value;
	}
	public static void totalrows() throws IOException {
		File f = new File("C:\\Users\\ARUNKUMAR\\eclipse-workspace\\Maven12PM\\ExcelSheet\\StudentDetail.xlsx");

		FileInputStream fin = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fin);
		Sheet s = w.getSheet("StudentDetail1");
		int prows = s.getPhysicalNumberOfRows();
		System.out.println("total no's rows:" + prows);
	}
	public static void totalcells() throws IOException {
		File f = new File("C:\\Users\\ARUNKUMAR\\eclipse-workspace\\Maven12PM\\ExcelSheet\\StudentDetail.xlsx");

		FileInputStream fin = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fin);
		Sheet s = w.getSheet("StudentDetail1");

		int a=0,b=0,c=0;
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row row = s.getRow(i);
			a++;

			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				b++;
			}
		}
		c=c+b;
		System.out.println(c);		
	}

	public static void newvalue(String value1,String value2, int a,int b,String value3) throws IOException {
		File file = new File(value1);
		FileInputStream fint = new FileInputStream(file);
		Workbook w = new XSSFWorkbook(fint);
		Sheet s = w.getSheet(value2);
		Row r = s.getRow(a);
		Cell c = r.createCell(b);
		c.setCellValue(value3);
		FileOutputStream fout = new FileOutputStream(file);
		w.write(fout);
		System.out.println("Done");

}
         public static void dropdown(WebElement a,String b) {
			Select s = new Select(a);
			s.selectByValue(b);
		}
         public static void dropdownid(WebElement a,int b) {
			Select s = new Select(a);
			s.selectByIndex(b);

		}

}







