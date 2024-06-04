package Other;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;
import java.util.concurrent.ThreadLocalRandom;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;

public class Sections_Testing_V1_0 {

	public static void main(String[] args) throws Exception {
		// Setup Drivers and Browser
		WebDriver wd;
		XSSFWorkbook wb = new XSSFWorkbook();
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		wd.manage().window().maximize();
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(15));
		WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(15));
		// ===== ALL URLs in Arrays =====
		String[] mainURLs = new String[25];
		String[] Versions = new String[25];
		

		mainURLs[0] = "https://east.merckmanuals.com/professional";
		Versions[0] = "EN US PV";
		mainURLs[1] = "https://www.msdmanuals.com/professional";
		Versions[1] = "EN PV";
		mainURLs[2] = "https://www.msdmanuals.com/es/professional";
		Versions[2] = "ES PV";
		mainURLs[3] = "https://www.msdmanuals.com/fr/professional";
		Versions[3] = "FR PV";
		mainURLs[4] = "https://www.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB";
		Versions[4] = "JA PV";
		mainURLs[5] = "https://www.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A";
		Versions[5] = "ZH PV";
		mainURLs[6] = "https://www.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9";
		Versions[6] = "RU PV";
		mainURLs[7] = "https://www.msdmanuals.com/pt/profissional";
		Versions[7] = "PT PV";
		mainURLs[8] = "https://www.msdmanuals.com/it/professionale";
		Versions[8] = "IT PV";
		mainURLs[9] = "https://www.msdmanuals.com/de/profi";
		Versions[9] = "DE PV";
		mainURLs[10] = "https://www.merckmanuals.com/home";
		Versions[10] = "EN US CV";
		mainURLs[11] = "https://www.msdmanuals.com/home";
		Versions[11] = "EN CV";
		mainURLs[12] = "https://www.msdmanuals.com/es/hogar";
		Versions[12] = "ES CV";
		mainURLs[13] = "https://www.msdmanuals.com/fr/accueil";
		Versions[13] = "FR CV";
		mainURLs[14] = "https://www.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0";
		Versions[14] = "JA CV";
		mainURLs[15] = "https://www.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5";
		Versions[15] = "ZH CV";
		mainURLs[16] = "https://www.msdmanuals.com/ko/%ED%99%88";
		Versions[16] = "KO CV";
		mainURLs[17] = "https://www.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0";
		Versions[17] = "RU CV";
		mainURLs[18] = "https://www.msdmanuals.com/pt/casa";
		Versions[18] = "PT CV";
		mainURLs[19] = "https://www.msdmanuals.com/it/casa";
		Versions[19] = "IT CV";
		mainURLs[20] = "https://www.msdmanuals.com/de/heim";
		Versions[20] = "DE CV";
		mainURLs[21] = "https://www.merckvetmanual.com";
		Versions[21] = "MM Vet";
		mainURLs[22] = "https://www.msdvetmanual.com";
		Versions[22] = "MSD Vet";
		mainURLs[23] = "https://www.msdmanuals.com/ar/home";
		Versions[23] = "AR CV";
		mainURLs[24] = "https://www.msdmanuals.com/vi/chuy%C3%AAn-gia";
		Versions[24] = "VI PV";

		// Main Loop
		for (int j = 0; j < 25; j++) {
			Thread.sleep(3000);
			XSSFSheet sheet = wb.createSheet(Versions[j]);
			Row rowHeading0 = sheet.createRow(0);
			rowHeading0.createCell(0).setCellValue("Test Page");
			rowHeading0.createCell(1).setCellValue("Page Title");
			rowHeading0.createCell(2).setCellValue("Element On Page");
			rowHeading0.createCell(3).setCellValue("Page URL");
			rowHeading0.createCell(4).setCellValue("Status");
			rowHeading0.createCell(5).setCellValue("Comments 1");
			rowHeading0.createCell(6).setCellValue("Comments 2");
			rowHeading0.createCell(7).setCellValue("Comments 3");
			try {
				// Navigate to version
				wd.get(mainURLs[j]);
				String CurrentVersion = Versions[j];
				System.out.println(Versions[j]);
				wd.navigate().refresh();

				// Handle modal windows
				if (CurrentVersion.contains("RU CV")) {
					Thread.sleep(3000);
					try {wd.findElement(By.xpath("//*[text()[contains(.,'Да')]]")).click();}catch(Exception e) {}
					Thread.sleep(1000);

				} else if (CurrentVersion.contains("ZH PV")) {
					Thread.sleep(3000);
					try {wd.findElement(By.xpath("//*[text()[contains(.,'我是专业医学人士')]]")).click();}catch(Exception e) {}
					Thread.sleep(1000);
				}

			} catch (Exception e) {
				System.out.println("N/A");
			}


			// START Sections 
			Row row2 = sheet.createRow(2);
			row2.createCell(0).setCellValue("Section Page");
			
			
			try {
				
				List<WebElement> Chapters = wd.findElements(By.xpath("//*[@class='mainmenu__item item1 topics col4 mainmenu__healthtopics']//div[@class='maintab__list--2col']/a"));
				int sizeChapters = Chapters.size();
				System.out.println("Size of Chapters: "+ sizeChapters);
				
						try {
							Thread.sleep(3000);
							Actions actions = new Actions(wd);
							WebElement target = wd.findElement(By.xpath("//*[@id=\"l-group--header-letterpine\"]/div[1]/div[3]/nav/div[2]"));
							actions.moveToElement(target).perform();
							Thread.sleep(3000);
							}catch(Exception ea) {
							System.out.println("Cannot Click on Medical Topics");
						}
						//Choose Random Chapter
						for (int i = 1; i < sizeChapters+1; i++) {
							wd.findElement(By.xpath("//*[@id=\"l-group--header-letterpine\"]/div[1]/div[3]/nav/div[2]/div[2]/div/div[1]/div[3]/div/a["+i+"]")).click();
							
							
							
							
							Actions actions = new Actions(wd);
							WebElement target = wd.findElement(By.xpath("//*[@id=\"l-group--header-letterpine\"]/div[1]/div[3]/nav/div[2]"));
							actions.moveToElement(target).perform();
							Thread.sleep(3000);
							
						}
						
						
				
					

			} catch (Exception e) {
				System.out.println("Failed");
				row2.createCell(4).setCellValue("Failed");
			}

			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			 DateFormat dateFormat = new SimpleDateFormat("MM.dd.yyyy");
			 
			 //get current date time with Date()
			 Date date = new Date();
			 
			 // Now format the date
			 String date1= dateFormat.format(date);
			 
			 // Print the Date
			 System.out.println("Current date and time is " +date1);
			FileOutputStream fout = new FileOutputStream("C:\\users\\gvahe\\Desktop\\Section_Testing_"+date1+".xlsx");
			wb.write(fout);
			fout.close();
			System.out.println("Written Successfully!");
			Thread.sleep(2000);
		}

		// Close excel and browser
		wb.close();
		wd.close();
		System.out.println("Browser Closed!");
	}

}
