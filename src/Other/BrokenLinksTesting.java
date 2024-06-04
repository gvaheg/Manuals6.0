package Other;

import java.io.FileOutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.asserts.SoftAssert;
public class BrokenLinksTesting {

	public static void main(String[] args) throws Exception {
		// Setup Drivers and Browser
		System.setProperty("webdriver.chrome.driver", "C:\\SeleniumDrivers\\chromedriver.exe");
		WebDriver wd = new ChromeDriver();
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(1));
		
		/*
		XSSFWorkbook wb = new XSSFWorkbook();
		// Get Date
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
		Date date = new Date();
		String date1 = dateFormat.format(date);
		// Excel Create Headings
		XSSFSheet sheet = wb.createSheet("Home Pages");
		Row rowHeading0 = sheet.createRow(0);
		Row rowHeading1 = sheet.createRow(1);
		rowHeading0.createCell(0).setCellValue("START DATE/TIME: " + date1);
		rowHeading1.createCell(0).setCellValue("Version");
		rowHeading1.createCell(1).setCellValue("Performance");
		rowHeading1.createCell(2).setCellValue("Accessibility");
		rowHeading1.createCell(3).setCellValue("Best Practices");
		rowHeading1.createCell(4).setCellValue("SEO");
		rowHeading1.createCell(5).setCellValue("First Contentful Paint");
		rowHeading1.createCell(6).setCellValue("Speed Index");
		rowHeading1.createCell(7).setCellValue("Largest Contentful Paint");
		rowHeading1.createCell(8).setCellValue("Time to Interactive");
		rowHeading1.createCell(9).setCellValue("Total Blocking Time");
		rowHeading1.createCell(10).setCellValue("Comulative Layout Shift");
		int rowNum = 2;
		*/
		
		wd.get("https://east.msdmanuals.com/es/professional/resourcespages/covid-19-useful-links");
		// Close Cookies
		try {
			WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
			AcceptCookies.click();
		} catch (Exception e) {
			System.out.println("Can't Close Cookies");
		}
		Thread.sleep(2000);
		
		// MAIN SCRIPT
		
		/*
		WebElement testURL = wd.findElement(By.xpath("/html/body/div[4]/main/section/section/main/div[7]/div[2]/p[3]/a"));
		System.out.println(testURL.getAttribute("href"));
		String URLs = testURL.getAttribute("href");
		
		HttpURLConnection conn = (HttpURLConnection)new URL(URLs).openConnection();
		conn.setRequestMethod("HEAD");
		conn.connect();
		int respCode = conn.getResponseCode();
		if(respCode>400)
		{
			System.out.println("Link with text "+testURL.getText()+" is broken.");
			Assert.assertTrue(false);
		}
		*/
		
		
		/*
		List<WebElement> footerLinks = wd.findElements(By.xpath(".//a[@class='footer__footnav--link']"));
		System.out.println(footerLinks);
		for (WebElement footerLink : footerLinks) {
			String URLs = footerLink.getAttribute("href");
			
			HttpURLConnection conn = (HttpURLConnection)new URL(URLs).openConnection();
			conn.setRequestMethod("HEAD");
			conn.connect();
			int respCode = conn.getResponseCode();
			if(respCode>400)
			{
				System.out.println("Link with text "+footerLink.getText()+" is broken.");
				Assert.assertTrue(false);
			}
			
			//System.out.println(URLs);
		}
		*/
		
		
		List<WebElement> footerLinks = wd.findElements(By.xpath(".//div[@class='row']//a"));
		//SoftAssert aAssert = new SoftAssert();
		for (WebElement footerLink : footerLinks) {
			String url = footerLink.getAttribute("href");
			HttpURLConnection conn = (HttpURLConnection)new URL(url).openConnection();
			conn.setRequestMethod("HEAD");
			conn.connect();
			int respCode = conn.getResponseCode();
			System.out.println(url+" code: "+respCode);
			/*
			if(respCode>400)
			{
				System.out.println("Link with text "+footerLink.getText()+" is broken.");
				aAssert.assertTrue(false);
			}
			*/
		}
		//aAssert.assertAll();
		wd.close();
		/*

		FileOutputStream fout = new FileOutputStream("D:\\TestResults\\BrokenLinks.xlsx");
		wb.write(fout);
		fout.close();
		System.out.println("Written Successfully!");
		
		// Close excel and browser
		//wb.close();
*/
		System.out.println("Browser Closed!");
	}

}
