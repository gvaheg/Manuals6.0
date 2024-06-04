package TestCases;

import static org.testng.Assert.assertTrue;

import java.io.File;
import java.util.List;
import java.util.Arrays;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.platform.commons.function.Try;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.io.File;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;

public class LandingFunctions {

	// START DRIVER
	static WebDriver driver;
	static XSSFWorkbook wb;


	// LANDING PAGE TESTING
	public void verifyLandingCases() throws Exception {
		driver.manage().window().maximize();
		CloseCookies();
		String currentURL = driver.getCurrentUrl();
		driver.get(currentURL);
		WebDriverWait wait20 = new WebDriverWait(driver, Duration.ofSeconds(20));
		// Function to get correct data from Excel
		String url = driver.getCurrentUrl();
		int ColumnNGet = 0, ColumnNWrite = 0, CoulumnNPass = 0;

		if (url.equals("https://uat102.merckmanuals.com/")) {
			ColumnNGet = 6;
			ColumnNWrite = 7;
			CoulumnNPass = 38;
		} 
		
		else if (url.equals("https://uat102.msdmanuals.com/")) {
			ColumnNGet = 8;
			ColumnNWrite = 9;
			CoulumnNPass = 39;
		} else if (url.equals("https://uat102.msdmanuals.com/de")) {
			ColumnNGet = 10;
			ColumnNWrite = 11;
			CoulumnNPass = 40;
		} else if (url.equals("https://uat102.msdmanuals.com/es")) {
			ColumnNGet = 12;
			ColumnNWrite = 13;
			CoulumnNPass = 41;
		}else if (url.equals("https://uat102.msdmanuals.com/fr")) {
			ColumnNGet = 14;
			ColumnNWrite = 15;
			CoulumnNPass = 42;
		}else if (url.equals("https://uat102.msdmanuals.com/it")) {
			ColumnNGet = 16;
			ColumnNWrite = 17;
			CoulumnNPass = 43;
		}else if (url.equals("https://uat102.msdmanuals.com/ja-jp")) {
			ColumnNGet = 18;
			ColumnNWrite = 19;
			CoulumnNPass = 44;
		}else if (url.equals("https://uat102.msdmanuals.com/ko")) {
			ColumnNGet = 20;
			ColumnNWrite = 21;
			CoulumnNPass = 45;
		}else if (url.equals("https://uat102.msdmanuals.com/pt")) {
			ColumnNGet = 22;
			ColumnNWrite = 23;
			CoulumnNPass = 46;
		}else if (url.equals("https://uat102.msdmanuals.com/ru")) {
			ColumnNGet = 24;
			ColumnNWrite = 25;
			CoulumnNPass = 47;
		}else if (url.equals("https://uat102.msdmanuals.cn/")) {
			ColumnNGet = 26;
			ColumnNWrite = 27;
			CoulumnNPass = 48;
		}else if (url.equals("https://uat102.msdmanuals.com/vi")) {
			ColumnNGet = 28;
			ColumnNWrite = 29;
			CoulumnNPass = 49;
		}else if (url.equals("https://uat102.msdmanuals.com/ar")) {
			ColumnNGet = 30;
			ColumnNWrite = 31;
			CoulumnNPass = 50;
		}else if (url.equals("https://uat102.msdmanuals.com/uk")) {
			ColumnNGet = 32;
			ColumnNWrite = 33;
			CoulumnNPass = 51;
		}else if (url.equals("https://uat102.merckvetmanual.com/")) {
			ColumnNGet = 34;
			ColumnNWrite = 35;
			CoulumnNPass = 52;
		}else if (url.equals("https://uat102.msdvetmanual.com/")) {
			ColumnNGet = 36;
			ColumnNWrite = 37;
			CoulumnNPass = 53;
		}
		
		else {
			System.out.println("ERROR IS WITH IF!");
		}

		int SaveNum = ColumnNGet - 2;
		// Load the Excel spreadsheet
		FileInputStream file = new FileInputStream("C:\\TestResults\\TestCases\\TestCases" + SaveNum + ".xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);

		// Create the output file and copy the data from the input file to it
		FileOutputStream output = new FileOutputStream("C:\\TestResults\\TestCases\\TestCases" + ColumnNGet + ".xlsx");
		XSSFWorkbook outputWorkbook = new XSSFWorkbook();
		XSSFSheet outputSheet = outputWorkbook.createSheet(sheet.getSheetName());

		for (Row row : sheet) {
			Row outputRow = outputSheet.createRow(row.getRowNum());
			for (Cell cell : row) {
				Cell outputCell = outputRow.createCell(cell.getColumnIndex());
				switch (cell.getCellType()) {
				case STRING:
					outputCell.setCellValue(cell.getStringCellValue());
					break;
				case NUMERIC:
					outputCell.setCellValue(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					outputCell.setCellValue(cell.getBooleanCellValue());
					break;
				case FORMULA:
					outputCell.setCellFormula(cell.getCellFormula());
					break;
				default:
					// Do nothing
					break;
				}
			}
		}

		// TC001 Header Logo Icon
		try {
			// Read data from the input file
			Row Row = sheet.getRow(2);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(2);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement logoicon = driver.findElement(By.xpath(".//img[@alt='logo']"));
			String SiteData = logoicon.getAttribute("src");
			newCellText.setCellValue(SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("Header Logo Icon: Pass");
				newCellPassFail.setCellValue("Pass");
				System.out.println("Wrote data: " + SiteData);
			} else {
				System.out.println("Header Title: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Header Logo Image: Fail");
		}

		// TC002 Header Logo Branding
		try {
			// Read data from the input file
			Row Row = sheet.getRow(3);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(3);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText1 = driver.findElement(By.xpath(".//span[@class='header__logo--word first']"));
			String Text1 = GetText1.getText();
			WebElement GetText2 = driver.findElement(By.xpath(".//span[@class='header__logo--word second']"));
			String Text2 = GetText2.getText();
			String SiteData = Text1 + " " + Text2;
			System.out.println(SiteData);
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Header Logo Branding: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Header Logo Branding: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Header Logo Branding");

		}

		// TC003 Header Title
		try {
			// Read data from the input file
			Row Row = sheet.getRow(4);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(4);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver.findElement(By.xpath(".//h1[@class='header__title']"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Header Title: Pass");
				newCellPassFail.setCellValue("Pass");
				System.out.println("Wrote data: " + SheetData);
			} else {
				System.out.println("Header Title: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Header Title");
		}

		// TC004 Header Description
		try {
			// Read data from the input file
			Row Row = sheet.getRow(5);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(5);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			WebElement GetText = driver.findElement(By.xpath(".//h2[@class='header__description']"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Header Description: Pass");
				newCellPassFail.setCellValue("Pass");
				System.out.println("Wrote data: " + SheetData);
			} else {
				System.out.println("Header Description: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Header Description");
		}

		// TC005 Landing Manuals Title
		try {
			// Read data from the input file
			Row Row = sheet.getRow(6);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(6);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			WebElement GetText = driver.findElement(By.xpath(".//h2[@class='landing__manuals--title']"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Landing Manuals Title: Pass");
				newCellPassFail.setCellValue("Pass");
				System.out.println("Wrote data: " + SheetData);
			} else {
				System.out.println("Landing Manuals Title: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Landing Manuals Title");
		}

		// TC006 Professional Box Img
		try {
			// Read data from the input file
			Row Row = sheet.getRow(7);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(7);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement img = driver.findElement(By.xpath(".//div[@class='version version--professional']/div[1]/img"));
			String SiteData = img.getAttribute("src");
			System.out.println(SiteData);
			newCellText.setCellValue(SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("Professional Box Img: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Professional Box Img: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Header Logo Image: Fail");
		}

		// TC007 Professional Box Branding
		try {
			// Read data from the input file
			Row Row = sheet.getRow(8);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(8);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver
					.findElement(By.xpath(".//div[@class='version version--professional']/div[1]/figure/a/div/div[1]"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Professional Box Branding: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Professional Box Branding: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Professional Box Branding");
		}

		// TC008 Professional Box Version Name
		try {
			// Read data from the input file
			Row Row = sheet.getRow(9);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(9);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver.findElement(
					By.xpath(".//div[@class='version version--professional']/div[1]/figure/a/div/figcaption/h3"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Professional Box Version Name: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Professional Box Version Name: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Professional Box Version Name");
		}

		// TC009 Professional Box Version Description
		try {
			// Read data from the input file
			Row Row = sheet.getRow(10);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(10);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver
					.findElement(By.xpath(".//div[@class='version version--professional']/div[1]/figure/a/div/h3[1]"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Professional Box Version Description: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Professional Box Version Description: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Professional Box Version Description");
		}

		// TC010 Consumer Box Img
		try {
			// Read data from the input file
			Row Row = sheet.getRow(11);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(11);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement img = driver.findElement(By.xpath(".//div[@class='version version--home']/div[1]/img"));
			String SiteData = img.getAttribute("src");
			System.out.println(SiteData);
			newCellText.setCellValue(SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("Consumer Box Img: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Consumer Box Img: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Header Logo Image: Fail");
		}

		// TC011 Consumer Box Branding
		try {
			// Read data from the input file
			Row Row = sheet.getRow(12);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(12);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver
					.findElement(By.xpath(".//div[@class='version version--home']/div[1]/figure/a/div/div[1]"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Consumer Box Branding: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Consumer Box Branding: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Consumer Box Branding");
		}

		// TC012 Consumer Box Version Name
		try {
			// Read data from the input file
			Row Row = sheet.getRow(13);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(13);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver
					.findElement(By.xpath(".//div[@class='version version--home']/div[1]/figure/a/div/figcaption/h3"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Consumer Box Version Name: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Consumer Box Version Name: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Consumer Box Version Name");
		}

		// TC013 Consumer Box Version Description
		try {
			// Read data from the input file
			Row Row = sheet.getRow(14);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(14);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver
					.findElement(By.xpath(".//div[@class='version version--home']/div[1]/figure/a/div/h3[1]"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Consumer Box Version Description: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Consumer Box Version Description: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Consumer Box Version Description");
		}

		// TC014 VET Box Img
		try {
			// Read data from the input file
			Row Row = sheet.getRow(15);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(15);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement img = driver.findElement(By.xpath(".//div[@class='version version--veterinary']/div[1]/img"));
				String SiteData = img.getAttribute("src");
				System.out.println(SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("VET Box Img: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("VET Box Img: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("VET Box Img: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
				
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("VET Box Img: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("VET Box Img: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
			
		} catch (Exception e) {
			System.out.println("Header Logo Image: Fail");
		}

		// TC015 VET Box Branding
		try {
			// Read data from the input file
			Row Row = sheet.getRow(16);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(16);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver
						.findElement(By.xpath(".//div[@class='version version--veterinary']/div[1]/figure/a/div/div[1]"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				if (SiteData.contains(SheetData)) {
					System.out.println("VET Box Branding: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("VET Box Branding: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("VET Box Branding: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("VET Box Branding: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("VET Box Branding: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}

			
		} catch (Exception e) {
			System.out.println("Cannot test VET Box Branding");
		}

		// TC016 VET Box Version Name
		try {
			// Read data from the input file
			Row Row = sheet.getRow(17);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(17);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver.findElement(
						By.xpath(".//div[@class='version version--veterinary']/div[1]/figure/a/div/figcaption/h3"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				if (SiteData.contains(SheetData)) {
					System.out.println("VET Box Version Name: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("VET Box Version Name: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("VET Box Version Name: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("VET Box Version Name: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("VET Box Version Name: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot test VET Box Version Name");
		}

		// TC017 VET Box Version Description
		try {
			// Read data from the input file
			Row Row = sheet.getRow(18);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(18);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver
						.findElement(By.xpath(".//div[@class='version version--veterinary']/div[1]/figure/a/div/h3[1]"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				if (SiteData.contains(SheetData)) {
					System.out.println("VET Box Version Description: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("VET Box Version Description: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("VET Box Version Description: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("VET Box Version Description: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("VET Box Version Description: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
			
		} catch (Exception e) {
			System.out.println("Cannot test VET Box Version Description");
		}

		// TC018 Mission Content Img
		try {
			// Read data from the input file
			Row Row = sheet.getRow(19);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(19);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement img = driver.findElement(By.xpath("//*[@id=\"mission\"]/div[2]/div[1]/img"));
			String SiteData = img.getAttribute("src");
			System.out.println(SiteData);
			newCellText.setCellValue(SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("Mission Content Img: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Mission Content Img: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Header Logo Image: Fail");
		}

		// TC019 Mission Content Title
		try {
			// Read data from the input file
			Row Row = sheet.getRow(20);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(20);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver.findElement(By.xpath(".//h2[@class='content__title']"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Mission Content Title: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Mission Content Title: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Mission Content Title");
		}
		// TC020 Mission Content Paragraph 1
		try {
			// Read data from the input file
			Row Row = sheet.getRow(21);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(21);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver.findElement(By.xpath("//*[@id=\"mission\"]/div[2]/div[2]/p[1]"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				if (SiteData.contains(SheetData)) {
					System.out.println("Mission Content P1: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Mission Content P1: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			} catch (Exception e) {
				System.out.println("Mission Content P1: Fail");
				newCellPassFail.setCellValue("Fail");
			}
			
		} catch (Exception e) {
			System.out.println("Cannot test Mission Content P1");
		}

		// TC021 Mission Content Paragraph 2
		try {
			// Read data from the input file
			Row Row = sheet.getRow(22);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(22);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver.findElement(By.xpath("//*[@id=\"mission\"]/div[2]/div[2]/p[2]"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Mission Content P1: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Mission Content P1: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Mission Content P1");
		}

		// TC022 Mission Content Paragraph 3
		try {
			// Read data from the input file
			Row Row = sheet.getRow(23);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(23);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver.findElement(By.xpath("//*[@id=\"mission\"]/div[2]/div[2]/p[3]"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Mission Content P1: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Mission Content P1: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test Mission Content P1");
		}

		// TC023 About Left Content Title
		try {
			// Read data from the input file
			Row Row = sheet.getRow(24);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(24);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver
					.findElement(By.xpath(".//div[@class='content__left']/h2[@class='content__title']"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("About Left Content Title: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("About Left Content Title: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot test About Left Content Title");
		}
		// TC024 About Left Content Description
		try {
			// Read data from the input file
			Row Row = sheet.getRow(25);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(25);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver
					.findElement(By.xpath(".//div[@class='content__left']/div[@class='content__paragraph']"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("About Left Content Description: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("About Left Content Description: Fail");
				newCellPassFail.setCellValue("Pass");
			}
		} catch (Exception e) {
			System.out.println("Cannot test About Left Content Description");
		}

		// TC025 About Left Content Logo 1
		try {
			// Read data from the input file
			Row Row = sheet.getRow(26);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(26);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement logoicon = driver.findElement(By.xpath(".//div[@class='content__left']/div/div/img"));
			String SiteData = logoicon.getAttribute("src");
			newCellText.setCellValue(SiteData);
			System.out.println("Site Data: " + SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("About Left Content Logo 1: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("About Left Content Logo 1: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("About Left Content Logo 1: Fail");
		}

		// TC026 About Left Content Logo 2 Icon
		try {
			// Read data from the input file
			Row Row = sheet.getRow(27);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(27);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement logoicon = driver
					.findElement(By.xpath(".//div[@class='content__picture--logosecond']/div/a/img[@alt='logo']"));
			String SiteData = logoicon.getAttribute("src");
			newCellText.setCellValue(SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("About Left Content Logo 2 Icon: Pass");
				newCellPassFail.setCellValue("Pass");
				System.out.println("Wrote data: " + SiteData);
			} else {
				System.out.println("About Left Content Logo 2 Icon: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("About Left Content Logo 2 Icon: Fail");
		}

		// TC027 About Left Content Logo 2 Branding
		try {
			// Read data from the input file
			Row Row = sheet.getRow(28);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(28);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText1 = driver.findElement(By.xpath(
					".//div[@class='content__picture--logosecond']/div/a/div/span[@class='header__logo--word first']"));
			String Text1 = GetText1.getText();
			WebElement GetText2 = driver.findElement(By.xpath(
					".//div[@class='content__picture--logosecond']/div/a/div/span[@class='header__logo--word second']"));
			String Text2 = GetText2.getText();
			String SiteData = Text1 + " " + Text2;
			System.out.println(SiteData);
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("About Left Content Logo 2 Branding: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("About Left Content Logo 2 Branding: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot Test About Left Content Logo 2 Branding");

		}

		// TC028 About Right Content Img
		try {
			// Read data from the input file
			Row Row = sheet.getRow(29);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(29);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement logoicon = driver
					.findElement(By.xpath(".//div[@class='content content--flexible']/div[2]/div/img"));
			String SiteData = logoicon.getAttribute("src");
			newCellText.setCellValue(SiteData);
			System.out.println("Site Data: " + SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("About Right Content Img: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("About Right Content Img: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("About Left Content Logo 1: Fail");
		}
		
		// TC029 Settings Button Dark Mode
		try {
			// Read data from the input file
			Row Row = sheet.getRow(30);
			Cell Cell = Row.getCell(ColumnNGet);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(30);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement settings = driver.findElement(By.xpath(".//button[@class='settings--opt-button']"));
				settings.click();
				// Find the element that represents the theme switcher
				WebElement themeSwitcher = driver.findElement(By.xpath(".//span[@class='slider round']"));
				// Store the current theme in a variable
				WebElement bodyElement = driver.findElement(By.tagName("body"));
				String oldTheme = bodyElement.getCssValue("background-color");
				System.out.println("Old Theme: " + oldTheme);
				// Click on the theme switcher
				themeSwitcher.click();
				// Store the new theme in a variable
				String newTheme = bodyElement.getCssValue("background-color");
				System.out.println("New Theme: " + newTheme);
				// Use an assertion to check if the new theme is different from the old theme
				Assert.assertNotEquals("The theme did not change after clicking the theme switcher", oldTheme,
						newTheme);
				System.out.println("Dark Mode: Pass");
				newCellPassFail.setCellValue("Pass");
			} catch (Exception e) {
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Can't Test Dark Light Mode");
		}
		// TC030 Settings Button Font Size
		try {
			// Read data from the input file
			Row Row = sheet.getRow(31);
			Cell Cell = Row.getCell(ColumnNGet);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(31);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				// Test Font Size 2
				WebElement font2 = driver.findElement(By.xpath(".//a[@class='font-size-2']"));
				font2.click();
				WebElement element2 = driver.findElement(By.tagName("p"));
				String fontSize2 = element2.getCssValue("font-size");
				// Test Font Size 3
				WebElement font3 = driver.findElement(By.xpath(".//a[@class='font-size-3']"));
				font3.click();
				WebElement element3 = driver.findElement(By.tagName("p"));
				String fontSize3 = element3.getCssValue("font-size");
				// Test Font Size 1
				WebElement font1 = driver.findElement(By.xpath(".//a[@class='font-size-1']"));
				font1.click();
				WebElement element1 = driver.findElement(By.tagName("p"));
				String fontSize1 = element1.getCssValue("font-size");
				//Result
				if (fontSize2.equals("15.75px") && fontSize3.equals("17.5px")&& fontSize1.equals("14px")){
				    newCellPassFail.setCellValue("Pass");
				} else {
				    System.out.println("Fail");
				    newCellPassFail.setCellValue("Fail");
				}
				
			} catch (Exception e) {
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("Can't Test Dark Light Mode");
		}

		
		// TC031 Content Link to GMK 1
		try {
			// Read data from the input file
			Row Row = sheet.getRow(32);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(32);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetLink = driver.findElement(By.xpath(".//p[@class='content__paragraph']/a"));
				GetLink.click();
				String SiteData = driver.getTitle();
				System.out.println("Page title: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Content Link to GMK 1: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Content Link to GMK 1: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				driver.navigate().back();
				Thread.sleep(2000);
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Content Link to GMK 1");
		}

		//TC032 Content Link to GMK 2
		try {
			// Read data from the input file
			Row Row = sheet.getRow(33);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(33);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetLink = driver.findElement(By.xpath(".//div[@class='content__paragraph']/p/a"));
				GetLink.click();
				Thread.sleep(2000);
				String SiteData = driver.getCurrentUrl();
				newCellText.setCellValue(SiteData);
				System.out.println("Page URL: " + SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Content Link to GMK 2: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Content Link to GMK 2: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
			driver.navigate().back();
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Cannot Test Content Link to GMK 2");
		}
		
		// TC033 Choose The Version PV
		try {
			// Read data from the input file
			Row Row = sheet.getRow(34);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(34);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetLink = driver.findElement(By.xpath(".//div[@class='version version--professional']/div/figure/a"));
			GetLink.click();
			String SiteData = driver.getTitle();
			System.out.println("Page title: " + SiteData);
			newCellText.setCellValue(SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("Choose The Version PV: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Choose The Version PV: Fail");
				newCellPassFail.setCellValue("Fail");
			}
			driver.navigate().back();
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Cannot Choose The Version PV");
		}
		// TC034 Choose The Version CV
		try {
			// Read data from the input file
			Row Row = sheet.getRow(35);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(35);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetLink = driver.findElement(By.xpath(".//div[@class='version version--home']/div/figure/a"));
			GetLink.click();
			String SiteData = driver.getTitle();
			System.out.println("Page title: " + SiteData);
			newCellText.setCellValue(SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("Choose The Version PV: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Choose The Version PV: Fail");
				newCellPassFail.setCellValue("Fail");
			}
			driver.navigate().back();
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Cannot Choose The Version PV");
		}
		// TC035 Choose The Version VET
		try {
			// Read data from the input file
			Row Row = sheet.getRow(36);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(36);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetLink = driver.findElement(By.xpath(".//div[@class='version version--veterinary']/div/figure/a"));
				GetLink.click();
				String SiteData = driver.getTitle();
				System.out.println("Page title: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Choose The Version VET: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Choose The Version VET: Fail");
					if (SheetData.equals("N/A")) {
						System.out.println("Choose The Version VET: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Choose The Version VET: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Choose The Version VET: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Choose The Version VET: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
			
			driver.navigate().back();
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Cannot Choose The Version PV");
		}
		
		//TC036 Footer Navigation
		try {
			// Read data from the input file
			Row Row = sheet.getRow(37);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(37);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			// Find the UL element with class "footer__list"
	        WebElement footerList = driver.findElement(By.className("footer__list"));
	        // Get all the <li> items under the UL element
	        List<WebElement> liItems = footerList.findElements(By.tagName("li"));
	        // Iterate over the <li> items and check whether each one is in the expected list
	        for (WebElement li : liItems) {
	            String SiteData = li.getText();
	            if (SheetData.contains(SiteData)) {
	                System.out.println(SiteData + " - Found");
	                newCellText.setCellValue(SheetData);
	                newCellPassFail.setCellValue("Pass");
	            } else {
	                System.out.println(SiteData + " - Not Found");
	                newCellText.setCellValue(SiteData + " - WRONG ITEM DISPLAYED");
	                newCellPassFail.setCellValue("Fail");
	                break;
	            }
	        }
		} catch (Exception e) {
			System.out.println("Cannot Test Footer");
		}

		//TC037 Footer Social Links
		try {
			Thread.sleep(1000);
			// Read data from the input file
			Row Row = sheet.getRow(38);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(38);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				// Find the UL element
				WebElement footerSocial = driver.findElement(By.xpath(".//div[@class='footer__social--desktop']/div/ul"));
				 // Find all <li> elements under UL
		        List<WebElement> liElements = footerSocial.findElements(By.xpath("li/a"));
		        for (WebElement li : liElements) {
		            String SiteData = li.getAttribute("href");
		            System.out.println(SiteData);
		            if (SheetData.contains(SiteData)) {
		                System.out.println(SiteData + " - Found");
		                newCellText.setCellValue(SheetData);
		                newCellPassFail.setCellValue("Pass");
		            } else {
		                System.out.println(SiteData + " - Not Found");
		                newCellText.setCellValue(SiteData);
		                if (SheetData.equals("N/A")) {
							System.out.println("Footer Social Links: N/A");
							newCellPassFail.setCellValue("N/A");
						} else {
							System.out.println("Footer Social Links: Fail");
							newCellPassFail.setCellValue("Fail");
						}
		                break;
		            }
		        }
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Footer Social Links: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Footer Social Links: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}

		} catch (Exception e) {
			System.out.println("Cannot Test Footer");
		}

		//TC038 Footer Accessibility Link
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(39);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(39);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				// Find the UL element
				WebElement footerAccessibility = driver.findElement(By.xpath(".//div[@class='footer-ea-desktop']/a"));
		            String SiteData = footerAccessibility.getAttribute("href");
		            System.out.println(SiteData);
		            if (SheetData.contains(SiteData)) {
		                System.out.println(SiteData + " - Found");
		                newCellText.setCellValue(SheetData);
		                newCellPassFail.setCellValue("Pass");
		            } else {
		                System.out.println(SiteData + " - Not Found");
		                newCellText.setCellValue(SiteData);
		                if (SheetData.equals("N/A")) {
							System.out.println("Footer Accessibility Link: N/A");
							newCellPassFail.setCellValue("N/A");
						} else {
							System.out.println("Footer Accessibility Link: Fail");
							newCellPassFail.setCellValue("Fail");
						}
		            }
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Footer Accessibility Link: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Footer Accessibility Link: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}

		} catch (Exception e) {
			System.out.println("Cannot Test Footer Accessibility Link");
		}
		
		
		//TC039 Footer Copyright Text
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(40);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(40);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			// Find the UL element
			WebElement footerAccessibility = driver.findElement(By.xpath(".//div[@class='footer__copyright__text']"));
	            String SiteData = footerAccessibility.getText();
	            System.out.println(SiteData);
	            if (SheetData.contains(SiteData)) {
	                System.out.println(SiteData + " - Found");
	                newCellText.setCellValue(SheetData);
	                newCellPassFail.setCellValue("Pass");
	            } else {
	                System.out.println(SiteData + " - Not Found");
	                newCellText.setCellValue(SiteData);
	                newCellPassFail.setCellValue("Fail");
	            }
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Copyright Text");
		}
		
		
		//TC040 Landing Footer Navigation - About
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(41);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			System.out.println("Next Cell Text: "+NextCellText);
			System.out.println("Cell Text: "+SheetDataClick);
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(41);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link About: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link About: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				
				
				
				driver.navigate().back();
				Thread.sleep(2000);
			} catch (Exception e) {
				System.out.println("Cannot Click on Footer About");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
			
		} catch (Exception e) {
			System.out.println("Cannot Test Footer About");
		}

		//TC041 Landing Footer Navigation - Disclaimer
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(42);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(42);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link About: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link About: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				driver.navigate().back();
				Thread.sleep(2000);
			} catch (Exception e) {
				System.out.println("Cannot Click on Footer Disclaimer");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Disclaimer");
		}
		
		//TC042 Landing Footer Navigation - Permissions
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(43);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(43);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link Permissions: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link Permissions: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				driver.navigate().back();
				Thread.sleep(2000);
			} catch (Exception e) {
				System.out.println("Cannot Click on Footer Permissions");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Permissions");
		}
		
		//TC043 Landing Footer Navigation - Privacy
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(44);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(44);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				// Get the window handles of all open tabs
				Set<String> allWindowHandles = driver.getWindowHandles();
				String mainWindowHandle = driver.getWindowHandle();
				// Loop through all window handles and switch to the new tab
				for (String windowHandle : allWindowHandles) {
					if (!windowHandle.equals(mainWindowHandle)) {
						driver.switchTo().window(windowHandle);
						break;
					}
				}
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link Privacy: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link Privacy: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				driver.close();
				driver.switchTo().window(mainWindowHandle);
				Thread.sleep(2000);
			} catch (Exception e) {
				System.out.println("Cannot Click on Footer Disclaimer");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Disclaimer");
		}
		
		//TC044 Landing Footer Navigation - Cookie Preferences
		try {
			Thread.sleep(1000);
			// Read data from the input file
			Row Row = sheet.getRow(45);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(45);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000); 
				WebElement SiteDataD = driver.findElement(By.xpath(".//h2[@id='ot-pc-title']"));
				String SiteData = SiteDataD.getText();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link Cookie Preferences: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link Cookie Preferences: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				WebElement CloseCookies1 = driver.findElement(By.xpath(".//*[@id=\"onetrust-pc-sdk\"]/div/div[3]/div[1]/button"));
				CloseCookies1.click();
				Thread.sleep(1000);
			} catch (Exception e) {
				System.out.println("Cannot Click on Cookie Preferences");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Cookie Preferences");
		}
		
		//TC045 Landing Footer Navigation - Terms of Use
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(46);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(46);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				// Get the window handles of all open tabs
				Set<String> allWindowHandles = driver.getWindowHandles();
				String mainWindowHandle = driver.getWindowHandle();
				// Loop through all window handles and switch to the new tab
				for (String windowHandle : allWindowHandles) {
					if (!windowHandle.equals(mainWindowHandle)) {
						driver.switchTo().window(windowHandle);
						break;
					}
				}
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link Terms of Use: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link Terms of Use: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				try {
					driver.close();
					driver.switchTo().window(mainWindowHandle);
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Window");
				}
				
			} catch (Exception e) {
				System.out.println("Cannot Click on Footer Terms of Use");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Terms of Use");
		}
		
		//TC046 Landing Footer Navigation - Licensing
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(47);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(47);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link Licensing: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link Licensing: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				driver.navigate().back();
				Thread.sleep(2000);
			} catch (Exception e) {
				System.out.println("Cannot Click on Licensing");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Licensing");
		}
		
		//TC047 Landing Footer Navigation - Contact Us
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(48);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(48);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link Contact Us: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link Contact Us: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				driver.navigate().back();
				Thread.sleep(2000);
			} catch (Exception e) {
				System.out.println("Cannot Click on Contact Us");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Contact Us");
		}
		
		//TC048 Landing Footer Navigation - Global Medical Knowledge
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(49);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(49);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link Global Medical Knowledge: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link Global Medical Knowledge: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				driver.navigate().back();
				Thread.sleep(2000);
			} catch (Exception e) {
				System.out.println("Cannot Click on Global Medical Knowledge");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Global Medical Knowledge");
		}
		
		
		//TC049 Landing Footer Navigation - Veterinary Manual
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(50);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(50);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				// Get the window handles of all open tabs
				Set<String> allWindowHandles = driver.getWindowHandles();
				String mainWindowHandle = driver.getWindowHandle();
				// Loop through all window handles and switch to the new tab
				for (String windowHandle : allWindowHandles) {
					if (!windowHandle.equals(mainWindowHandle)) {
						driver.switchTo().window(windowHandle);
						break;
					}
				}
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link Veterinary Manual: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link Veterinary Manual: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				try {
					driver.close();
					driver.switchTo().window(mainWindowHandle);
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Window");
				}
			} catch (Exception e) {
				System.out.println("Cannot Click on Veterinary Manual");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Veterinary Manual");
		}
		
		
		//TC050 Landing Footer Navigation - Print Editions
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(51);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(51);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link Print Editions: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link Print Editions: Fail");
					newCellPassFail.setCellValue("N/A");
				}
				driver.navigate().back();
				Thread.sleep(2000);
			} catch (Exception e) {
				System.out.println("Cannot Click on Print Editions");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Print Editions");
		}
		
		
		//TC051 Landing Footer Navigation - Mobile Apps
		try {
			Thread.sleep(2000);
			// Read data from the input file
			Row Row = sheet.getRow(52);
			Cell CelltoClick = Row.getCell(ColumnNGet);
			Cell NextCellText = Row.getCell(ColumnNWrite);
			System.out.println("Cell Text: "+CelltoClick);
			System.out.println("Next Cell Text: "+NextCellText);
			String SheetDataClick = CelltoClick.getStringCellValue().toLowerCase();
			String SheetDataCompare = NextCellText.getStringCellValue();
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(52);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FooterLink = driver.findElement(By.xpath(".//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'"+SheetDataClick+"')]"));
				FooterLink.click();
				Thread.sleep(1000);
				// Get the window handles of all open tabs
				Set<String> allWindowHandles = driver.getWindowHandles();
				String mainWindowHandle = driver.getWindowHandle();
				// Loop through all window handles and switch to the new tab
				for (String windowHandle : allWindowHandles) {
					if (!windowHandle.equals(mainWindowHandle)) {
						driver.switchTo().window(windowHandle);
						break;
					}
				}
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetDataCompare)) {
					System.out.println("Footer Link Mobile Apps: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Link Mobile Apps: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				try {
					driver.close();
					driver.switchTo().window(mainWindowHandle);
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Window");
				}
			} catch (Exception e) {
				System.out.println("Cannot Click on Mobile Apps");
				if (SheetDataClick.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Footer Mobile Apps");
		}
		/*
		// Language Bar Testing
		try {
			String CurrentURL = driver.getCurrentUrl();
			if (CurrentURL.contains("merck")) {
				MMLanguageSwitch();
				OtherLanguages();
			} else {
				System.out.println("Cannot Test MSD VERSIONS YET");
			}
		} catch (Exception e) {
			System.out.println("Cannot Test Language Bar");
		}
		
*/
        
		// Save the changes to the output file
		workbook.close();
		file.close();
		outputWorkbook.write(output);
		output.close();
		outputWorkbook.close();
		System.out.println("Wrote Data Successfully!");

	}

	
	// LANGUAGE BAR! SWITCH LANGUAGES
		public void MMLanguageSwitch() {
		try {
			// Go to ES-US Version
			WebElement languageSwitcher = driver.findElement(By.id("ui-id-1-button"));
			// Store the current language in a variable
			WebElement currentLanguageElement = driver.findElement(By.xpath(".//option[@selected='\"selected\"']"));
			String currentLanguage = currentLanguageElement.getAttribute("innerText");
			System.out.println("Current Language: " + currentLanguage);
			languageSwitcher.click();
			Thread.sleep(2000);
			// Click on ES-US version and Verify ES-US is displayed
			WebElement element = driver.findElement(By.id("ui-id-3"));
			element.click();
			String linkik = driver.getCurrentUrl();
			System.out.println(linkik);
			if (linkik.equals("https://uat102.merckmanuals.com/es-us")) {
				System.out.println("Language Switched to ES-US: Pass");
			} else {
				System.out.println("Language Switched to ES-US: Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot locate switcher");
		}
		try {
			// Go back to EN-US version
			WebElement languageSwitcher1 = driver.findElement(By.id("ui-id-1-button"));
			WebElement currentLanguageElement1 = driver.findElement(By.xpath(".//option[@selected='\"selected\"']"));
			String currentLanguage1 = currentLanguageElement1.getAttribute("innerText");
			System.out.println("Current Language: " + currentLanguage1);
			languageSwitcher1.click();
			Thread.sleep(2000);
			WebElement element1 = driver.findElement(By.id("ui-id-2"));
			element1.click();
			String linkik1 = driver.getCurrentUrl();
			System.out.println(linkik1);
			if (linkik1.equals("https://uat102.merckmanuals.com/")) {
				System.out.println("Language Switched to EN-US: Pass");
			} else {
				System.out.println("Language Switched to EN-US: Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot locate switcher");
		}
	}
	
	
	// LANGUAGE BAR! OTHER LANGUAGES
	public void OtherLanguages() {
		try {
			// Go to Other Languages
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(40));
			WebElement languageSwitcher2 = driver.findElement(By.id("ui-id-1-button"));
			WebElement currentLanguageElement2 = driver.findElement(By.xpath(".//option[@selected='\"selected\"']"));
			String currentLanguage2 = currentLanguageElement2.getAttribute("innerText");
			System.out.println("Current Language: " + currentLanguage2);
			languageSwitcher2.click();
			Thread.sleep(2000);
			WebElement element2 = driver.findElement(By.id("ui-id-4"));
			element2.click();
			// Wait for the new tab to open
			wait.until(ExpectedConditions.numberOfWindowsToBe(2));
			// Get the window handles of all open tabs
			Set<String> allWindowHandles = driver.getWindowHandles();
			String mainWindowHandle = driver.getWindowHandle();
			// Loop through all window handles and switch to the new tab
			for (String windowHandle : allWindowHandles) {
				if (!windowHandle.equals(mainWindowHandle)) {
					driver.switchTo().window(windowHandle);
					break;
				}
			}
			try {
				WebElement AcceptCookies = driver.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
				AcceptCookies.click();
			} catch (Exception e) {
				System.out.println("Can't Close Cookies");
			}
			Thread.sleep(1000);
			String linkik2 = driver.getCurrentUrl();
			System.out.println(linkik2);
			if (linkik2.equals("https://uat102.msdmanuals.com/?langselector=1")) {
				System.out.println("Language Selector Page: Pass");
			} else {
				System.out.println("Language Selector Page: Fail");
			}
		} catch (Exception e) {
			System.out.println("Cannot locate switcher");
		}
	}

	

	////////////////////////// REUSABLE METHODS
	
	
	// TEMPLATE FOR TEST CASES
	public void verifyTEMPCHANGE() throws Exception {

		driver.manage().window().maximize();
		CloseCookies();
		String currentURL = driver.getCurrentUrl();
		driver.get(currentURL);

		// Function to get correct data from Excel
		String url = driver.getCurrentUrl();
		int ColumnNGet = 0, ColumnNWrite = 0, CoulumnNPass = 0;

		if (url.contains("https://uat102.merckmanuals.com")) {
			ColumnNGet = 6;
			ColumnNWrite = 7;
			CoulumnNPass = 38;
		} else if (url.contains("https://uat102.msdmanuals.com")) {
			ColumnNGet = 8;
			ColumnNWrite = 9;
			CoulumnNPass = 39;
		} else {
			System.out.println("ERROR IS WITH IF!");
		}

		int SaveNum = ColumnNGet - 2;
		// Load the Excel spreadsheet
		FileInputStream file = new FileInputStream("C:\\TestResults\\TestCases\\TestCases" + SaveNum + ".xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);

		// Create the output file and copy the data from the input file to it
		FileOutputStream output = new FileOutputStream("C:\\TestResults\\TestCases\\TestCases" + ColumnNGet + ".xlsx");
		XSSFWorkbook outputWorkbook = new XSSFWorkbook();
		XSSFSheet outputSheet = outputWorkbook.createSheet(sheet.getSheetName());

		for (Row row : sheet) {
			Row outputRow = outputSheet.createRow(row.getRowNum());
			for (Cell cell : row) {
				Cell outputCell = outputRow.createCell(cell.getColumnIndex());
				switch (cell.getCellType()) {
				case STRING:
					outputCell.setCellValue(cell.getStringCellValue());
					break;
				case NUMERIC:
					outputCell.setCellValue(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					outputCell.setCellValue(cell.getBooleanCellValue());
					break;
				case FORMULA:
					outputCell.setCellFormula(cell.getCellFormula());
					break;
				default:
					// Do nothing
					break;
				}
			}
		}

		
		
		
		//CODE GOES HERE!!!!!!!!!!!!!!
		
		
		
		
		
		// Save the changes to the output file
		workbook.close();
		file.close();
		outputWorkbook.write(output);
		output.close();
		outputWorkbook.close();
		System.out.println("Wrote Data Successfully!");

	}
	
	// Get Date
	public void getDate() throws Exception {
		// Create object of SimpleDateFormat class and decide the format
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
		// get current date time with Date()
		Date date = new Date();
		// Now format the date
		String date1 = dateFormat.format(date);
		// Print the Date
		System.out.println("Current date/time: " + date1);
	}
	
	// Manage Windows
	public void ManageWindows() {
		// driver.manage().window().maximize();
		driver.manage().window().setSize(new Dimension(400, 820));
		driver.manage().window().setPosition(new Point(780, 0));
		// driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(1));
		Actions action = new Actions(driver);
	}

	// Close Cookies
	public void CloseCookies() throws InterruptedException {
		try {
			WebElement AcceptCookies = driver.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
			AcceptCookies.click();
		} catch (Exception e) {
			System.out.println("Can't Close Cookies");
		}
		Thread.sleep(1000);
	}

	// Confirm page opens
	public void confirmPage(WebDriverWait wait50, WebDriverWait wait20) {
		try {
			wait20.until(
					ExpectedConditions.visibilityOfElementLocated(By.xpath(".//section[1]/div/div/ul/li[1]/div/a")));
		} catch (Exception e) {
			System.out.println("Refreshing the page!");
			driver.navigate().refresh();
			wait20.until(
					ExpectedConditions.visibilityOfElementLocated(By.xpath(".//section[1]/div/div/ul/li[1]/div/a")));
		}
	}


	// Open Each Resource Item
	public void openResource(Row rowN, int i, int s) throws InterruptedException {

		try {
			WebElement Section = driver
					.findElement(By.xpath(".//section[" + s + "]/div[@class='resources__container--header']/h3"));
			rowN.createCell(0).setCellValue(Section.getAttribute("innerText"));
			System.out.println("Section: " + Section.getAttribute("innerText"));
			rowN.createCell(1).setCellValue(i);

		} catch (Exception e) {
			System.out.println("Can't Find Section");
		}

		try {
			WebElement elementToClick = driver
					.findElement(By.xpath(".//section[" + s + "]/div/div/ul/li[" + i + "]/div/a"));
			try {
				JavascriptExecutor js = (JavascriptExecutor) driver;
				WebElement ScrollTo = driver.findElement(By.xpath(".//section[" + s + "]/div/div/ul/li[" + i + "]"));
				js.executeScript("arguments[0].scrollIntoView(true);", ScrollTo);
				js.executeScript("javascript:window.scrollBy(0,-100)");
			} catch (Exception e) {
				System.out.println("Can't SCROLLL!!!");
			}
			elementToClick.click();

		} catch (Exception e) {
			Thread.sleep(2000);
			System.out.println("Can't CLICK ON ELEMENT");
			try {
				WebElement elementToClick = driver
						.findElement(By.xpath(".//section[" + s + "]/div/div/ul/li[" + i + "]/div/a"));
				elementToClick.click();
			} catch (Exception ex) {
				Thread.sleep(2000);
				WebElement elementToClick = driver
						.findElement(By.xpath(".//section[" + s + "]/div/div/ul/li[" + i + "]/div/a"));
				elementToClick.click();
				System.out.println("Can't CLICK ON ELEMENT 2nd time");
			}
		}

		try {
			WebElement Title = driver.findElement(By.xpath(".//div[@class='multimedia__description--title']"));
			System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
			// Write ID and Title
			rowN.createCell(1).setCellValue(i);
			rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
		} catch (Exception e) {
			System.out.println("Can't get the title. Reloading...");

			try {
				((JavascriptExecutor) driver).executeScript("window.open()");
				Thread.sleep(1000);
				ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
				driver.close();
				Thread.sleep(2000);
				driver.switchTo().window(tabs.get(1));
				CloseCookies();
				WebElement elementToClick = driver
						.findElement(By.xpath(".//section[" + s + "]/div/div/ul/li[" + i + "]/div/a"));
				elementToClick.click();
				WebElement Title = driver.findElement(By.xpath(".//div[@class='multimedia__description--title']"));
				System.out.println("Title " + i + ": " + Title.getAttribute("innerText"));
				rowN.createCell(1).setCellValue(i);
				rowN.createCell(2).setCellValue(Title.getAttribute("innerText"));
			} catch (Exception ex) {
				System.out.println("Can't get title second time...");
			}

		}

	}

	// Get Credits and Descriptions
	public void getCredits(Row rowN) {
		// Get & write Credits
		try {
			WebElement credits = driver.findElement(By.xpath(".//div[@class='multimedia__description--credits']"));
			System.out.println("Credits: " + credits.getAttribute("innerText"));
			rowN.createCell(3).setCellValue(credits.getAttribute("innerText"));
		} catch (Exception e) {
			System.out.println("");
			rowN.createCell(3).setCellValue("");
		}
		// Get & write Description
		try {
			WebElement description = driver
					.findElement(By.xpath(".//div/div[@class='multimedia__description--content']/div"));
			System.out.println("Description: " + description.getAttribute("innerText"));
			rowN.createCell(4).setCellValue(description.getAttribute("innerText"));
		} catch (Exception e) {
			System.out.println("Description: ");
			rowN.createCell(4).setCellValue("");
		}

	}

	// GET FILE NAME AND VERIFY IMAGE
	public void getFileName(Row rowN, WebDriverWait wait10) {
		try {
			WebElement FileName = driver.findElement(By.xpath(".//div[@class='multimedia__image--wrapper']/img"));
			URL FileURL = new URL(FileName.getAttribute("src"));
			System.out.println("Src attribute is: " + FileURL);
			System.out.println("path = " + FileURL.getPath());
			rowN.createCell(9).setCellValue(FileName.getAttribute("src"));
			rowN.createCell(10).setCellValue(FileURL.getPath());
			String ImageURL = FileName.getAttribute("src");
			// VERIFY LINKS ARE ACTIVE
			try {
				((JavascriptExecutor) driver).executeScript("window.open()");
				Thread.sleep(1000);
				ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
				try {
					driver.switchTo().window(tabs.get(1));
					driver.get(ImageURL);
				} catch (Exception e) {
					Thread.sleep(3000);
					driver.switchTo().window(tabs.get(1));
					driver.get(ImageURL);
				}
				try {
					try {
						wait10.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/img")));
					} catch (Exception e) {
						System.out.println("Refreshing the page!");
						driver.navigate().refresh();
						wait10.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/img")));
					}
					WebElement ImagePresent = driver.findElement(By.xpath("/html/body/img"));
					System.out.println("Image: FOUND");
					rowN.createCell(11).setCellValue("FOUND");
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Image: NOT FOUND");
					rowN.createCell(11).setCellValue("NOT FOUND");
				}
				driver.close();
				driver.switchTo().window(tabs.get(0));
			} catch (Exception e) {
				Thread.sleep(2000);
				System.out.println("CANNOT VERIFY IMAGE");
				rowN.createCell(11).setCellValue("CANNOT VERIFY");
			}
		} catch (Exception e) {
			System.out.println("CANNOT GET FILE NAME");
			rowN.createCell(9).setCellValue("No File PATH");
			rowN.createCell(10).setCellValue("No File URL");
			rowN.createCell(11).setCellValue("No Image");
		}
	}

	// Get Location
	public void getLocation(Row rowN, WebDriverWait wait10, WebDriverWait wait20, WebDriverWait wait30, int i,
			Row rowHeading4, int HeadN, int s) throws IOException, InterruptedException {
		try {
			wait10.until(ExpectedConditions
					.visibilityOfElementLocated(By.xpath(".//div/a[@class='multimedia__description--location']")));
			WebElement location = driver.findElement(By.xpath(".//div/a[@class='multimedia__description--location']"));
			System.out.println("Location " + i + ": " + location.getAttribute("innerText"));
			// Excel Values
			rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
			rowN.createCell(6).setCellValue("Pass");
			String linkUrl = location.getAttribute("href");
			rowN.createCell(7).setCellValue(linkUrl);
			System.out.println(linkUrl);

			// VERIFY LINKS ARE ACTIVE
			try {
				((JavascriptExecutor) driver).executeScript("window.open()");
				Thread.sleep(1000);
				ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
				try {
					driver.switchTo().window(tabs.get(1));
					driver.get(linkUrl);
				} catch (Exception e) {
					Thread.sleep(3000);
					driver.switchTo().window(tabs.get(1));
					driver.get(linkUrl);
				}

				try {
					CloseCookies();

					try {
						wait30.until(ExpectedConditions
								.visibilityOfElementLocated(By.xpath(".//div[@class='topic__headings']/div/div/h1")));
					} catch (Exception e) {
						System.out.println("Refreshing the page!");
						driver.navigate().refresh();
						wait30.until(ExpectedConditions
								.visibilityOfElementLocated(By.xpath(".//div[@class='topic__headings']/div/div/h1")));
					}
					WebElement TopicTitle = driver.findElement(By.xpath(".//div[@class='topic__headings']/div/div/h1"));
					System.out.println("Location Topic Title: " + TopicTitle.getText());
					rowN.createCell(8).setCellValue("OK");
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Location Topic Title: NOT FOUND");
					rowN.createCell(8).setCellValue("Not Found");
				}
				int HrefNum = HeadN++;
				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='de']"));
					System.out.println("DE HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='es']"));
					System.out.println("ES HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='fr']"));
					System.out.println("FR HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='it']"));
					System.out.println("IT HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='ja-JP']"));
					System.out.println("JA HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='pt']"));
					System.out.println("PT HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='ru']"));
					System.out.println("RU HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}

				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='zh']"));
					System.out.println("CN HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}

				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='vi']"));
					System.out.println("VI HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}
				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='ar']"));
					System.out.println("AR HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}

				try {
					WebElement HrefLink = driver.findElement(By.xpath(".//link[@hreflang='uk']"));
					System.out.println("UK HREF Link: " + HrefLink.getAttribute("href"));
					rowN.createCell(HrefNum++).setCellValue(HrefLink.getAttribute("href"));
				} catch (Exception e) {
					System.out.println("Couldn't get HREF");
					rowN.createCell(HrefNum++).setCellValue("N/A");
				}

				// END HREFs
				driver.close();
				driver.switchTo().window(tabs.get(0));
			} catch (Exception e) {
				Thread.sleep(2000);
				System.out.println("URL: CANNOT VERIFY");
				driver.close();
				rowN.createCell(8).setCellValue("CANNOT VERIFY");
			}
		} catch (Exception e) {
			System.out.println("Location " + i + ": Fail");
			Thread.sleep(3000);
			rowN.createCell(1).setCellValue(i);
			rowN.createCell(6).setCellValue("Fail");
		} // END get location, verify links

	}

	// Close Popup
	public void ClosePopup() throws InterruptedException {

		try {
			// Close Popup
			driver.switchTo().defaultContent();
			driver.findElement(By.xpath(".//div[@class='modaldialog__header--close modaldialog__header--element']"))
					.click();
		} catch (Exception e) {
			System.out.println("Cannot Close Popup!");
			Thread.sleep(2000);
			ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
			driver.switchTo().window(tabs.get(0));
			Thread.sleep(2000);
			driver.switchTo().defaultContent();
			driver.findElement(By.xpath(".//div[@class='modaldialog__header--close modaldialog__header--element']"))
					.click();
		}
	}

	// HANDLING ERRORS
	public void ErrorMain(String currentURL, Exception e) throws Exception {
		Thread.sleep(3000);
		((JavascriptExecutor) driver).executeScript("window.open()");
		Thread.sleep(2000);
		ArrayList<String> tabs1 = new ArrayList<String>(driver.getWindowHandles());
		Thread.sleep(2000);
		driver.close();
		driver.switchTo().window(tabs1.get(1));
		driver.get(currentURL);
		Thread.sleep(5000);
		Thread.sleep(2000);
		CloseCookies();
	}



}
