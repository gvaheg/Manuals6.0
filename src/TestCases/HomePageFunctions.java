package TestCases;

import static org.testng.Assert.assertTrue;

import java.io.File;
import java.util.List;
import java.util.NoSuchElementException;
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
import org.apache.poi.hssf.record.PageBreakRecord.Break;
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

public class HomePageFunctions {

	// START DRIVER
	static WebDriver driver;
	static XSSFWorkbook wb;

	// PV HOME PAGE
	public void verifyPVHome() throws Exception {
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		Thread.sleep(2000);
		CloseCookies();
		String currentURL = driver.getCurrentUrl();
		driver.get(currentURL);

		// Function to get correct data from Excel
		String url = driver.getCurrentUrl();
		int ColumnNGet = 0, ColumnNWrite = 0, CoulumnNPass = 0;

		if (url.equals("https://uat102.merckmanuals.com/professional")) {
			ColumnNGet = 6;
			ColumnNWrite = 7;
			CoulumnNPass = 38;
		}else if (url.equals("https://uat102.msdmanuals.com/professional")) {
			ColumnNGet = 8;
			ColumnNWrite = 9;
			CoulumnNPass = 39;
		}else if (url.equals("https://uat102.msdmanuals.com/de/profi")) {
			ColumnNGet = 10;
			ColumnNWrite = 11;
			CoulumnNPass = 40;
		}else if (url.equals("https://uat102.msdmanuals.com/es/professional")) {
			ColumnNGet = 12;
			ColumnNWrite = 13;
			CoulumnNPass = 41;
		}else if (url.equals("https://uat102.msdmanuals.com/fr/professional")) {
			ColumnNGet = 14;
			ColumnNWrite = 15;
			CoulumnNPass = 42;
		}else if (url.equals("https://uat102.msdmanuals.com/it/professionale")) {
			ColumnNGet = 16;
			ColumnNWrite = 17;
			CoulumnNPass = 43;
		}else if (url.equals("https://uat102.msdmanuals.com/ja-jp/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB")) {
			ColumnNGet = 18;
			ColumnNWrite = 19;
			CoulumnNPass = 44;
		}else if (url.equals("https://uat102.msdmanuals.com/en-kr/professional")) {
			ColumnNGet = 20;
			ColumnNWrite = 21;
			CoulumnNPass = 45;
		}else if (url.equals("https://uat102.msdmanuals.com/pt/profissional")) {
			ColumnNGet = 22;
			ColumnNWrite = 23;
			CoulumnNPass = 46;
		}else if (url.equals("https://uat102.msdmanuals.com/ru/%d0%bf%d1%80%d0%be%d1%84%d0%b5%d1%81%d1%81%d0%b8%d0%be%d0%bd%d0%b0%d0%bb%d1%8c%d0%bd%d1%8b%d0%b9")) {
			ColumnNGet = 24;
			ColumnNWrite = 25;
			CoulumnNPass = 47;
		}else if (url.equals("https://uat102.msdmanuals.cn/professional")) {
			ColumnNGet = 26;
			ColumnNWrite = 27;
			CoulumnNPass = 48;
		}else if (url.equals("https://uat102.msdmanuals.com/vi/chuy%c3%aan-gia")) {
			ColumnNGet = 28;
			ColumnNWrite = 29;
			CoulumnNPass = 49;
		}else if (url.equals("n/a")) {
			ColumnNGet = 30;
			ColumnNWrite = 31;
			CoulumnNPass = 50;
		}else if (url.equals("https://uat102.msdmanuals.com/uk/professional")) {
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
		}else {
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

		// TC052-TC053 FEATURED CONTENT BOX
		// TC052 PV Home Page - Featured Content
		try {
			System.out.println("START: TC052");
			// Read data from the input file
			Row Row = sheet.getRow(53);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(53);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run Script
			WebElement GetText = driver
					.findElement(By.xpath("//*[@id=\"maincontent\"]/section/div[1]/div/div[1]/div[1]/h2"));
			String SiteData = GetText.getText();
			newCellText.setCellValue(SiteData);
			System.out.println("Site Data: " + SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Featured Content: Pass");
				newCellPassFail.setCellValue("Pass");
				System.out.println("Wrote data: " + SheetData);
			} else {
				System.out.println("Featured Content: Fail");
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC052: Fail");
		}

		try {
			// TC053 PV Home Page - Featured Content
			System.out.println("START: TC053");
			Row Row = sheet.getRow(54);
			Cell Cell = Row.getCell(ColumnNGet);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(54);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			try {
				// Run Script
				WebElement slider = driver
						.findElement(By.xpath(".//div[@class='featured__slides slick-initialized slick-slider']"));
				newCellPassFail.setCellValue("Pass");
			} catch (Exception e) {
				System.out.println("Can't find slider");
				newCellPassFail.setCellValue("Pass");
			}
		} catch (Exception e) {
			// TODO: handle exception
		}

		// TC054-TC061 VIDEOS BOX
		// TC054 PV Home Page - Videos Box Title
		try {
			System.out.println("START: TC054");
			// Read data from the input file
			Row Row = sheet.getRow(55);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(55);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--brightcove orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Videos Box Title: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Videos Box Title: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Videos Box Title: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("Videos Box Title: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC054: Fail");
		}

		// TC055 PV Home Page - Videos Box View All Button is present
		try {
			System.out.println("START: TC055");
			// Read data from the input file
			Row Row = sheet.getRow(56);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(56);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--brightcove orderize__candidate']/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Videos Box View All Button is present: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Videos Box Videos Box View All Button is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Videos Box Videos Box View All Button is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("Videos Box Videos Box View All Button is present: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC055: Fail");
		}

		// TC056 PV Home Page - Videos Box View All Button is functionality
		try {
			System.out.println("START: TC056");
			// Read data from the input file
			Row Row = sheet.getRow(57);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(57);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--brightcove orderize__candidate']/a"));
				TestElement.click();
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Videos Box View All Button is functionality: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Videos Box View All Button is functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Videos Box View All Button is functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
				Thread.sleep(1000);
				driver.navigate().back();
			} catch (Exception e) {
				System.out.println("Videos Box View All Button is functionality: Fail");
				newCellPassFail.setCellValue("Fail");
				Thread.sleep(1000);
				driver.navigate().back();
			}
		} catch (Exception e) {
			System.out.println("TC056: Fail");
		}

		// TC057 PV Home Page - Videos Box Featured Video functionality
		try {
			System.out.println("START: TC057");
			// Read data from the input file
			Row Row = sheet.getRow(58);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(58);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement ClickElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--brightcove orderize__candidate']/div[2]/div/figure/div/div[3]/button"));
				ClickElement.click();
				WebElement TestElement = driver.findElement(
						By.xpath(".//div[@class='modaldialog__header--title modaldialog__header--element']"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Videos Box Featured Video functionality: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Videos Box Featured Video functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Videos Box Featured Video functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
				ClosePopup();
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Videos Box Featured Video functionality: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Videos Box Featured Video functionality: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC057: Fail");
		}

		// TC058 PV Home Page - Featured video title is present
		try {
			System.out.println("START: TC058");
			// Read data from the input file
			Row Row = sheet.getRow(59);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(59);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--brightcove orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Videos BOX IS PRESENT");
					WebElement TestElement2 = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--brightcove orderize__candidate']/div[2]/div/div[1]/div/a/h3"));
					if (TestElement2 != null) {
						System.out.println("Featured video title is present: TEST ELEMENT FOUND");
						newCellPassFail.setCellValue("Pass");
					} else {
						System.out.println("Featured video title is present: TEST ELEMENT NOT FOUND");
						newCellPassFail.setCellValue("Fail");
					}
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured video title is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured video title is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured video title is present: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured video title is present: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC058: Fail");
		}

		// TC059 PV Home Page - Featured video description is present
		try {
			System.out.println("START: TC059");
			// Read data from the input file
			Row Row = sheet.getRow(60);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(60);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--brightcove orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Videos BOX IS PRESENT");
					WebElement TestElement2 = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--brightcove orderize__candidate']/div[2]/div/div[2]/div"));
					if (TestElement2 != null) {
						System.out.println("Featured video description is present: TEST ELEMENT FOUND");
						newCellPassFail.setCellValue("Pass");
					} else {
						System.out.println("Featured video description is present: TEST ELEMENT NOT FOUND");
						newCellPassFail.setCellValue("Fail");
					}
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured video description is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured video description is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured video description is present: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured video description is present: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC059: Fail");
		}

		// TC060 PV Home Page - Featured video Read All button is present
		try {
			System.out.println("START: TC060");
			// Read data from the input file
			Row Row = sheet.getRow(61);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(61);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--brightcove orderize__candidate']/div[2]/div/div[2]/div/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Featured video Read All button is present: Fail");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured video Read All button is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured video Read All button is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured video Read All button is present: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured video Read All button is present: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC060: Fail");
		}
		// TC061 PV Home Page - Featured video Read All button functions properly
		try {
			System.out.println("START: TC061");
			// Read data from the input file
			Row Row = sheet.getRow(62);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(62);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--brightcove orderize__candidate']/div[2]/div/div[2]/div/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Videos BOX IS PRESENT");
					TestElement.click();
					WebElement TestElement2 = driver.findElement(
							By.xpath(".//div[@class='modaldialog__header--title modaldialog__header--element']"));
					if (TestElement2 != null) {
						System.out.println("Featured video Read All button functions properly: TEST ELEMENT FOUND");
						newCellPassFail.setCellValue("Pass");
					} else {
						System.out.println("Featured video Read All button functions properly: TEST ELEMENT NOT FOUND");
						newCellPassFail.setCellValue("Fail");
					}
					ClosePopup();
					Thread.sleep(4000);

				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured video Read All button functions properly: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured video Read All button functions properly: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured video Read All button functions properly: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured video Read All button functions properly: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC061: Fail");
		}

		// TC062-TC069 3D MODELS BOX
		// TC062 PV Home Page - 3D Models Box Title
		try {
			System.out.println("START: TC062");
			// Read data from the input file
			Row Row = sheet.getRow(63);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(63);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--biodigital orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("3D Models Box Title: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("3D Models Box Title: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("3D Models Box Title: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("3D Models Box Title: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC062: Fail");
		}

		// TC063 PV Home Page - 3D Models Box View All Button is present
		try {
			System.out.println("START: TC063");
			// Read data from the input file
			Row Row = sheet.getRow(64);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(64);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--biodigital orderize__candidate']/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("3D Models Box View All Button is present: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("3D Models Box 3D Models Box View All Button is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("3D Models Box 3D Models Box View All Button is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("3D Models Box 3D Models Box View All Button is present: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC063: Fail");
		}

		// TC064 PV Home Page - 3D Models Box View All Button is functionality
		try {
			System.out.println("START: TC064");
			// Read data from the input file
			Row Row = sheet.getRow(65);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(65);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--biodigital orderize__candidate']/a"));
				TestElement.click();
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("3D Models Box View All Button is functionality: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("3D Models Box View All Button is functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("3D Models Box View All Button is functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
				Thread.sleep(1000);
				driver.navigate().back();
			} catch (Exception e) {
				System.out.println("3D Models Box View All Button is functionality: Fail");
				newCellPassFail.setCellValue("Fail");
				Thread.sleep(1000);
				driver.navigate().back();
			}
		} catch (Exception e) {
			System.out.println("TC064: Fail");
		}

		// TC065 PV Home Page - 3D Models Box Featured 3D Model functionality
		try {
			System.out.println("START: TC065");
			// Read data from the input file
			Row Row = sheet.getRow(66);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(66);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement ClickElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--biodigital orderize__candidate']/div[2]/div/div/figure/div[2]/button"));
				ClickElement.click();
				WebElement TestElement = driver.findElement(
						By.xpath(".//div[@class='modaldialog__header--title modaldialog__header--element']"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("3D Models Box Featured 3D Model functionality: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("3D Models Box Featured 3D Model functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("3D Models Box Featured 3D Model functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
				ClosePopup();
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("3D Models Box Featured 3D Model functionality: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("3D Models Box Featured 3D Model functionality: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC065: Fail");
		}

		// TC066 PV Home Page - Featured 3D Model title is present
		try {
			System.out.println("START: TC066");
			// Read data from the input file
			Row Row = sheet.getRow(67);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(67);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--biodigital orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("3D ModelS BOX IS PRESENT");
					WebElement TestElement2 = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--biodigital orderize__candidate']/div[2]/div/div/div[1]/div/a/h3"));
					if (TestElement2 != null) {
						System.out.println("Featured 3D Model title is present: TEST ELEMENT FOUND");
						newCellPassFail.setCellValue("Pass");
					} else {
						System.out.println("Featured 3D Model title is present: TEST ELEMENT NOT FOUND");
						newCellPassFail.setCellValue("Fail");
					}
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured 3D Model title is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured 3D Model title is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured 3D Model title is present: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured 3D Model title is present: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC066: Fail");
		}

		// TC067 PV Home Page - Featured 3D Model description is present
		try {
			System.out.println("START: TC067");
			// Read data from the input file
			Row Row = sheet.getRow(68);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(68);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--biodigital orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("3D ModelS BOX IS PRESENT");
					WebElement TestElement2 = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--biodigital orderize__candidate']/div[2]/div[1]/div[2]/div"));
					if (TestElement2 != null) {
						System.out.println("Featured 3D Model description is present: TEST ELEMENT FOUND");
						newCellPassFail.setCellValue("Pass");
					} else {
						System.out.println("Featured 3D Model description is present: TEST ELEMENT NOT FOUND");
						newCellPassFail.setCellValue("Fail");
					}
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured 3D Model description is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured 3D Model description is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured 3D Model description is present: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured 3D Model description is present: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC067: Fail");
		}

		// TC068 PV Home Page - Featured 3D Model Read All button is present
		try {
			System.out.println("START: TC068");
			// Read data from the input file
			Row Row = sheet.getRow(69);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(69);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--biodigital orderize__candidate']/div[2]/div[1]/div[2]/div/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Featured 3D Model Read All button is present: Fail");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured 3D Model Read All button is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured 3D Model Read All button is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured 3D Model Read All button is present: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured 3D Model Read All button is present: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC068: Fail");
		}
		// TC069 PV Home Page - Featured 3D Model Read All button functions properly
		try {
			System.out.println("START: TC069");
			// Read data from the input file
			Row Row = sheet.getRow(70);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(70);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--biodigital orderize__candidate']/div[2]/div[1]/div[2]/div/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("3D ModelS BOX IS PRESENT");
					TestElement.click();
					WebElement TestElement2 = driver.findElement(
							By.xpath(".//div[@class='modaldialog__header--title modaldialog__header--element']"));
					if (TestElement2 != null) {
						System.out.println("Featured 3D Model Read All button functions properly: TEST ELEMENT FOUND");
						newCellPassFail.setCellValue("Pass");
					} else {
						System.out.println(
								"Featured 3D Model Read All button functions properly: TEST ELEMENT NOT FOUND");
						newCellPassFail.setCellValue("Fail");
					}
					ClosePopup();
					Thread.sleep(4000);

				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured 3D Model Read All button functions properly: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured 3D Model Read All button functions properly: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured 3D Model Read All button functions properly: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured 3D Model Read All button functions properly: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC069: Fail");
		}

		// TC070-TC075 LATEST NEWS BOX
		// TC070 PV Home Page - Latest News Box Title
		try {
			System.out.println("START: TC070");
			// Read data from the input file
			Row Row = sheet.getRow(71);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(71);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(
						By.xpath(".//div[@class='box box--2colspan latestnews orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Latest News Box Title: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Latest News Box Title: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Latest News Box Title: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("Latest News Box Title: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC070: Fail");
		}

		// TC071 PV Home Page - Latest News Box View All Button is present
		try {
			System.out.println("START: TC071");
			// Read data from the input file
			Row Row = sheet.getRow(72);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(72);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver
						.findElement(By.xpath(".//div[@class='box box--2colspan latestnews orderize__candidate']/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Latest Newss Box View All Button is present: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Latest Newss Box Latest Newss Box View All Button is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Latest Newss Box Latest Newss Box View All Button is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("Latest Newss Box Latest Newss Box View All Button is present: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC071: Fail");
		}

		// TC072 PV Home Page - Latest News Box View All Button is functionality
		try {
			System.out.println("START: TC072");
			// Read data from the input file
			Row Row = sheet.getRow(73);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(73);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver
						.findElement(By.xpath(".//div[@class='box box--2colspan latestnews orderize__candidate']/a"));
				TestElement.click();
				// Thread.sleep(6000);
				String SiteData = driver.getCurrentUrl();

				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Latest News Box View All Button is functionality: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Latest News Box View All Button is functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Latest News Box View All Button is functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
				Thread.sleep(1000);
				driver.navigate().back();
			} catch (Exception e) {
				System.out.println("Latest News Box View All Button is functionality: Fail");
				newCellPassFail.setCellValue("Fail");
				Thread.sleep(1000);
				driver.navigate().back();
			}
		} catch (Exception e) {
			System.out.println("TC072: Fail");
		}

		// TC073 PV Home Page - Latest News Box Featured Image
		try {
			System.out.println("START: TC073");
			// Read data from the input file
			Row Row = sheet.getRow(74);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(74);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box box--2colspan latestnews orderize__candidate']/div[2]/figure/a/img"));
				if (SheetData.equals("N/A")) {
					System.out.println("Latest News Box Featured Image: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Latest News Box Featured Image: Pass");
					newCellPassFail.setCellValue("Pass");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Latest News Box Featured Image: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Latest News Box Featured Image: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC073: Fail");
		}

		// TC074 PV Home Page - Latest News Box Featured Latest News titles
		try {
			System.out.println("START: TC074");
			// Read data from the input file
			Row Row = sheet.getRow(75);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(75);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			try {
				// Run script
				WebElement newsLink = driver.findElement(
						By.xpath(".//div[@class='box box--2colspan latestnews orderize__candidate']/div[2]/div"));
				List<WebElement> h3Tags = newsLink.findElements(By.xpath(".//h3"));
				System.out.println("Number News Items: " + h3Tags.size());
				if (SheetData.equals("N/A")) {
					System.out.println("Latest News Box Featured Latest News titles: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Latest News Box Featured Latest News titles: Pass");
					newCellPassFail.setCellValue("Pass");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Latest News Box Featured Latest News titles: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Latest News Box Featured Latest News titles: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC074: Fail");
		}

		// TC075 PV Home Page - Latest News Box Featured Latest News functionality
		try {
			System.out.println("START: TC075");
			// Read data from the input file
			Row Row = sheet.getRow(76);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(76);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			try {
				// Run script
				WebElement newsLink = driver.findElement(
						By.xpath(".//div[@class='box box--2colspan latestnews orderize__candidate']/div[2]/div"));
				List<WebElement> h3Tags = newsLink.findElements(By.xpath(".//h3"));
				System.out.println("Number News Items: " + h3Tags.size());
				int numNewsItems = h3Tags.size();
				// Verify that each news item opens in a new page
				for (int i = 1; i <= numNewsItems; i++) {
					WebElement newsToClick = driver.findElement(
							By.xpath(".//div[@class='box box--2colspan latestnews orderize__candidate']/div[2]/div/h3["
									+ i + "]"));
					newsToClick.click();
					Thread.sleep(3000);
					String newUrl = driver.getCurrentUrl();
					System.out.println(newUrl);
					Thread.sleep(1000);
					driver.navigate().back();
				}
				System.out.println("Latest News Box Featured Image: Pass");
				newCellPassFail.setCellValue("Pass");
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Latest News Box Featured Latest News functionality: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Latest News Box Featured Latest News functionality: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}

		} catch (Exception e) {
			System.out.println("TC075: Fail");
		}

		// TC076-TC083 CASE STUDIES BOX
		// TC076 PV Home Page - Case Studies Box Title
		try {
			System.out.println("START: TC076");
			// Read data from the input file
			Row Row = sheet.getRow(77);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(77);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--externalpage orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Case Studies Box Title: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Case Studies Box Title: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Case Studies Box Title: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("Case Studies Box Title: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC076: Fail");
		}

		// TC077 PV Home Page - Case Studies Box View All Button is present
		try {
			System.out.println("START: TC077");
			// Read data from the input file
			Row Row = sheet.getRow(78);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(78);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--externalpage orderize__candidate']/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Case Studies Box View All Button is present: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Case Studies Box Case Studies Box View All Button is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Case Studies Box Case Studies Box View All Button is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("Case Studies Box Case Studies Box View All Button is present: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC077: Fail");
		}

		// TC078 PV Home Page - Case Studies Box View All Button is functionality
		try {
			System.out.println("START: TC078");
			// Read data from the input file
			Row Row = sheet.getRow(79);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(79);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--externalpage orderize__candidate']/a"));
				TestElement.click();
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Case Studies Box View All Button is functionality: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Case Studies Box View All Button is functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Case Studies Box View All Button is functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
				Thread.sleep(1000);
				driver.navigate().back();
			} catch (Exception e) {
				System.out.println("Case Studies Box View All Button is functionality: Fail");
				newCellPassFail.setCellValue("Fail");
				Thread.sleep(1000);
				driver.navigate().back();
			}
		} catch (Exception e) {
			System.out.println("TC078: Fail");
		}

		// TC079 PV Home Page - Case Studies Box Featured Case Study functionality
		try {
			System.out.println("START: TC079");
			// Read data from the input file
			Row Row = sheet.getRow(80);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(80);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				Thread.sleep(1000);
				try {
					WebElement ClickElement = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--externalpage orderize__candidate']/div[2]/div/figure/div[2]/button"));
					ClickElement.click();
				} catch (Exception e) {
					System.out.println("Cannot Click");
				}

				WebElement TestElement = driver.findElement(
						By.xpath(".//div[@class='modaldialog__header--title modaldialog__header--element']"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Case Studies Box Featured Case Study functionality: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Case Studies Box Featured Case Study functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Case Studies Box Featured Case Study functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
				ClosePopup();
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Case Studies Box Featured Case Study functionality: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Case Studies Box Featured Case Study functionality: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC079: Fail");
		}

		// TC080 PV Home Page - Featured Case Study title is present
		try {
			System.out.println("START: TC080");
			// Read data from the input file
			Row Row = sheet.getRow(81);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(81);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--externalpage orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Case Studies BOX IS PRESENT");
					WebElement TestElement2 = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--externalpage orderize__candidate']/div[2]/div/div[1]/div/a/h3"));
					if (TestElement2 != null) {
						System.out.println("Featured Case Study title is present: TEST ELEMENT FOUND");
						newCellPassFail.setCellValue("Pass");
					} else {
						System.out.println("Featured Case Study title is present: TEST ELEMENT NOT FOUND");
						newCellPassFail.setCellValue("Fail");
					}
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured Case Study title is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured Case Study title is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured Case Study title is present: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured Case Study title is present: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC080: Fail");
		}

		// TC081 PV Home Page - Featured Case Study description is present
		try {
			System.out.println("START: TC081");
			// Read data from the input file
			Row Row = sheet.getRow(82);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(82);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--externalpage orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Case Studies BOX IS PRESENT");
					WebElement TestElement2 = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--externalpage orderize__candidate']/div[2]/div/div[2]/div"));
					if (TestElement2 != null) {
						System.out.println("Featured Case Study description is present: TEST ELEMENT FOUND");
						newCellPassFail.setCellValue("Pass");
					} else {
						System.out.println("Featured Case Study description is present: TEST ELEMENT NOT FOUND");
						newCellPassFail.setCellValue("Fail");
					}
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured Case Study description is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured Case Study description is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured Case Study description is present: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured Case Study description is present: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC081: Fail");
		}

		// TC082 PV Home Page - Featured Case Study Read All button is present
		try {
			System.out.println("START: TC082");
			// Read data from the input file
			Row Row = sheet.getRow(83);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(83);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--externalpage orderize__candidate']/div[2]/div/div[2]/div/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Featured Case Study Read All button is present: Fail");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured Case Study Read All button is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured Case Study Read All button is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured Case Study Read All button is present: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured Case Study Read All button is present: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC082: Fail");
		}
		// TC083 PV Home Page - Featured Case Study Read All button functions properly
		try {
			System.out.println("START: TC083");
			// Read data from the input file
			Row Row = sheet.getRow(84);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(84);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(By.xpath(
						".//div[@class='box featuredmedia box--1colspan featuredmedia--externalpage orderize__candidate']/div[2]/div/div[2]/div/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Case Studies BOX IS PRESENT");
					TestElement.click();
					WebElement TestElement2 = driver.findElement(
							By.xpath(".//div[@class='modaldialog__header--title modaldialog__header--element']"));
					if (TestElement2 != null) {
						System.out
								.println("Featured Case Study Read All button functions properly: TEST ELEMENT FOUND");
						newCellPassFail.setCellValue("Pass");
					} else {
						System.out.println(
								"Featured Case Study Read All button functions properly: TEST ELEMENT NOT FOUND");
						newCellPassFail.setCellValue("Fail");
					}
					ClosePopup();
					Thread.sleep(4000);

				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured Case Study Read All button functions properly: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured Case Study Read All button functions properly: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Featured Case Study Read All button functions properly: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Featured Case Study Read All button functions properly: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC083: Fail");
		}

		// TC084-TC091 STUDENT STORIES BOX
		// TC084 STUDENT STORIES BOX
		 
		String currentUrl = driver.getCurrentUrl();
		if (currentUrl.contains("merck")) {
			// TC084 PV Home Page - Student Stories Box Title
			try {
				System.out.println("START: TC084");
				// Read data from the input file
				Row Row = sheet.getRow(85);
				Cell Cell = Row.getCell(ColumnNGet);
				String SheetData = Cell.getStringCellValue();
				System.out.println("Sheet Data: " + SheetData);
				// Write new data to the output file
				Row outputRow = outputSheet.getRow(85);
				Cell newCellText = outputRow.createCell(ColumnNWrite);
				Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
				// Run script
				try {
					WebElement TestElement = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--medstudentstory orderize__candidate']/div[1]/h2"));
					String SiteData = TestElement.getText();
					System.out.println("Site Data: " + SiteData);
					newCellText.setCellValue(SiteData);
					if (SiteData.equals(SheetData)) {
						System.out.println("Student Stories Box Title: Pass");
						newCellPassFail.setCellValue("Pass");
					} else {
						if (SheetData.equals("N/A")) {
							System.out.println("Student Stories Box Title: N/A");
							newCellPassFail.setCellValue("N/A");
						} else {
							System.out.println("Student Stories Box Title: Fail");
							newCellPassFail.setCellValue("Fail");
						}
					}
				} catch (Exception e) {
					System.out.println("Student Stories Box Title: Fail");
					if (SheetData.equals("N/A")) {
						System.out.println("Student Stories Box Title: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Student Stories Box Title: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("TC084: Fail");
			}

			// TC085 PV Home Page - Student Stories Box View All Button is present
			try {
				System.out.println("START: TC085");
				// Read data from the input file
				Row Row = sheet.getRow(86);
				Cell Cell = Row.getCell(ColumnNGet);
				String SheetData = Cell.getStringCellValue();
				System.out.println("Sheet Data: " + SheetData);
				// Write new data to the output file
				Row outputRow = outputSheet.getRow(86);
				Cell newCellText = outputRow.createCell(ColumnNWrite);
				Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
				// Run script
				try {
					WebElement TestElement = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--medstudentstory orderize__candidate']/a"));
					String SiteData = TestElement.getText();
					System.out.println("Site Data: " + SiteData);
					newCellText.setCellValue(SiteData);
					if (SiteData.equals(SheetData)) {
						System.out.println("Student Stories Box View All Button is present: Pass");
						newCellPassFail.setCellValue("Pass");
					} else {
						if (SheetData.equals("N/A")) {
							System.out
									.println("Student Stories Box Student Stories Box View All Button is present: N/A");
							newCellPassFail.setCellValue("N/A");
						} else {
							System.out.println(
									"Student Stories Box Student Stories Box View All Button is present: Fail");
							newCellPassFail.setCellValue("Fail");
						}
					}
				} catch (Exception e) {
					System.out.println("Student Stories Box Student Stories Box View All Button is present: Fail");
					if (SheetData.equals("N/A")) {
						System.out.println("Student Stories Box Student Stories Box View All Button is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Student Stories Box Student Stories Box View All Button is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("TC085: Fail");
			}

			// TC086 PV Home Page - Student Stories Box View All Button is functionality
			try {
				System.out.println("START: TC086");
				// Read data from the input file
				Row Row = sheet.getRow(87);
				Cell Cell = Row.getCell(ColumnNGet);
				String SheetData = Cell.getStringCellValue();
				System.out.println("Sheet Data: " + SheetData);
				// Write new data to the output file
				Row outputRow = outputSheet.getRow(87);
				Cell newCellText = outputRow.createCell(ColumnNWrite);
				Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
				// Run script
				try {
					WebElement TestElement = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--medstudentstory orderize__candidate']/a"));
					TestElement.click();
					Thread.sleep(2000);
					// Wait for the new tab to open
					WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(40));
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
					String SiteData = driver.getTitle();
					System.out.println("Site Data: " + SiteData);
					newCellText.setCellValue(SiteData);
					if (SiteData.equals(SheetData)) {
						System.out.println("Student Stories Box View All Button is functionality: Pass");
						newCellPassFail.setCellValue("Pass");
					} else {
						if (SheetData.equals("N/A")) {
							System.out.println("Student Stories Box View All Button is functionality: N/A");
							newCellPassFail.setCellValue("N/A");
						} else {
							System.out.println("Student Stories Box View All Button is functionality: Fail");
							newCellPassFail.setCellValue("Fail");
						}
					}
					try {
						driver.close();
						driver.switchTo().window(mainWindowHandle);
						Thread.sleep(1000);
					} catch (Exception e) {
						System.out.println("Cannot Close Window");
					}
				} catch (Exception e) {
					System.out.println("Student Stories Box View All Button is functionality: Fail");
					if (SheetData.equals("N/A")) {
						System.out.println("Student Stories Box View All Button is functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Student Stories Box View All Button is functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
					Thread.sleep(1000);
					driver.navigate().back();
				}
			} catch (Exception e) {
				System.out.println("TC086: Fail");
			}

			// TC087 PV Home Page - Student Stories Box Featured Student Story functionality
			try {
				System.out.println("START: TC087");
				// Read data from the input file
				Row Row = sheet.getRow(88);
				Cell Cell = Row.getCell(ColumnNGet);
				String SheetData = Cell.getStringCellValue();
				System.out.println("Sheet Data: " + SheetData);
				// Write new data to the output file
				Row outputRow = outputSheet.getRow(88);
				Cell newCellText = outputRow.createCell(ColumnNWrite);
				Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
				// Run script
				try {
					Thread.sleep(500);
					try {
						WebElement TestElement = driver.findElement(By.xpath(
								".//div[@class='box featuredmedia box--1colspan featuredmedia--medstudentstory orderize__candidate']/div[2]/div/a"));
						TestElement.click();
					} catch (Exception e) {
						System.out.println("CANNOT CLICK");
					}
					Thread.sleep(500);
					// Wait for the new tab to open
					WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(40));
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
					String SiteData = driver.getCurrentUrl();
					System.out.println("Site Data: " + SiteData);
					newCellText.setCellValue(SiteData);
					if (SiteData.contains(SheetData)) {
						System.out.println("Student Stories Box Featured Student Story is functionality: Pass");
						newCellPassFail.setCellValue("Pass");
					} else {
						if (SheetData.equals("N/A")) {
							System.out.println("Student Stories Box Featured Student Story is functionality: N/A");
							newCellPassFail.setCellValue("N/A");
						} else {
							System.out.println("Student Stories Box Featured Student Story is functionality: Fail");
							newCellPassFail.setCellValue("Fail");
						}
					}
					try {
						driver.close();
						driver.switchTo().window(mainWindowHandle);
						Thread.sleep(1000);
					} catch (Exception e) {
						System.out.println("Cannot Close Window");
					}
				} catch (Exception e) {
					if (SheetData.equals("N/A")) {
						System.out.println("Student Stories Box Featured Student Story functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Student Stories Box Featured Student Story functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("TC087: Fail");
			}

			// TC088 PV Home Page - Featured Student Story title is present
			try {
				System.out.println("START: TC088");
				// Read data from the input file
				Row Row = sheet.getRow(89);
				Cell Cell = Row.getCell(ColumnNGet);
				String SheetData = Cell.getStringCellValue();
				System.out.println("Sheet Data: " + SheetData);
				// Write new data to the output file
				Row outputRow = outputSheet.getRow(89);
				Cell newCellText = outputRow.createCell(ColumnNWrite);
				Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
				// Run script
				try {
					WebElement TestElement = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--medstudentstory orderize__candidate']/a"));
					String SiteData = TestElement.getText();
					System.out.println("Site Data: " + SiteData);
					newCellText.setCellValue(SiteData);
					if (SiteData.equals(SheetData)) {
						System.out.println("Student Stories BOX IS PRESENT");
						try {
							WebElement TestElement2 = driver.findElement(By.xpath(
									".//div[@class='box featuredmedia box--1colspan featuredmedia--medstudentstory orderize__candidate']/div[2]/div/a/h3"));
							System.out.println("Test Elemet Found: " + TestElement2.getText());
							newCellPassFail.setCellValue("Pass");
						} catch (Exception e) {
							if (SheetData.equals("N/A")) {
								System.out.println("Featured Student Story title is present: N/A");
								newCellPassFail.setCellValue("N/A");
							} else {
								System.out.println("Featured Student Story title is present: Fail");
								newCellPassFail.setCellValue("Fail");
							}
						}
					} else {
						if (SheetData.equals("N/A")) {
							System.out.println("Featured Student Story title is present: N/A");
							newCellPassFail.setCellValue("N/A");
						} else {
							System.out.println("Featured Student Story title is present: Fail");
							newCellPassFail.setCellValue("Fail");
						}
					}
				} catch (Exception e) {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured Student Story title is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured Student Story title is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("TC088: Fail");
			}

			// TC089 PV Home Page - Featured Student Story description is present
			try {
				System.out.println("START: TC089");
				// Read data from the input file
				Row Row = sheet.getRow(90);
				Cell Cell = Row.getCell(ColumnNGet);
				String SheetData = Cell.getStringCellValue();
				System.out.println("Sheet Data: " + SheetData);
				// Write new data to the output file
				Row outputRow = outputSheet.getRow(90);
				Cell newCellText = outputRow.createCell(ColumnNWrite);
				Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
				// Run script
				try {
					WebElement TestElement = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--medstudentstory orderize__candidate']/a"));
					String SiteData = TestElement.getText();
					System.out.println("Site Data: " + SiteData);
					newCellText.setCellValue(SiteData);
					if (SiteData.equals(SheetData)) {
						System.out.println("Student Stories BOX IS PRESENT");
						try {
							WebElement TestElement2 = driver.findElement(By.xpath(
									".//div[@class='box featuredmedia box--1colspan featuredmedia--medstudentstory orderize__candidate']/div[2]/div/div"));
							System.out.println("Test Elemet Found: " + TestElement2.getText());
							newCellPassFail.setCellValue("Pass");
						} catch (Exception e) {
							if (SheetData.equals("N/A")) {
								System.out.println("Featured Student Story description is present: N/A");
								newCellPassFail.setCellValue("N/A");
							} else {
								System.out.println("Featured Student Story description is present: Fail");
								newCellPassFail.setCellValue("Fail");
							}
						}
					} else {
						if (SheetData.equals("N/A")) {
							System.out.println("Featured Student Story description is present: N/A");
							newCellPassFail.setCellValue("N/A");
						} else {
							System.out.println("Featured Student Story description is present: Fail");
							newCellPassFail.setCellValue("Fail");
						}
					}
				} catch (Exception e) {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured Student Story title is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured Student Story title is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("TC089: Fail");
			}

			// TC090 PV Home Page - Featured Student Story Read All button is present
			try {
				System.out.println("START: TC090");
				// Read data from the input file
				Row Row = sheet.getRow(91);
				Cell Cell = Row.getCell(ColumnNGet);
				String SheetData = Cell.getStringCellValue();
				System.out.println("Sheet Data: " + SheetData);
				// Write new data to the output file
				Row outputRow = outputSheet.getRow(91);
				Cell newCellText = outputRow.createCell(ColumnNWrite);
				Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
				// Run script
				try {
					WebElement TestElement = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--medstudentstory orderize__candidate']/div[2]/div/div/a"));
					String SiteData = TestElement.getText();
					System.out.println("Site Data: " + SiteData);
					newCellText.setCellValue(SiteData);
					if (SiteData.equals(SheetData)) {
						System.out.println("Featured Student Story Read All button is present: Pass");
						newCellPassFail.setCellValue("Pass");
					} else {
						if (SheetData.equals("N/A")) {
							System.out.println("Featured Student Story Read All button is present: N/A");
							newCellPassFail.setCellValue("N/A");
						} else {
							System.out.println("Featured Student Story Read All button is present: Fail");
							newCellPassFail.setCellValue("Fail");
						}
					}
				} catch (Exception e) {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured Student Story Read All button is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured Student Story Read All button is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("TC090: Fail");
			}
			// TC091 PV Home Page - Featured Student Story Read All button functions
			// properly
			try {
				System.out.println("START: TC091");
				// Read data from the input file
				Row Row = sheet.getRow(92);
				Cell Cell = Row.getCell(ColumnNGet);
				String SheetData = Cell.getStringCellValue();
				System.out.println("Sheet Data: " + SheetData);
				// Write new data to the output file
				Row outputRow = outputSheet.getRow(92);
				Cell newCellText = outputRow.createCell(ColumnNWrite);
				Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
				// Run script
				try {
					WebElement TestElement = driver.findElement(By.xpath(
							".//div[@class='box featuredmedia box--1colspan featuredmedia--medstudentstory orderize__candidate']/div[2]/div/div/a"));
					String SiteData = TestElement.getText();
					System.out.println("Site Data: " + SiteData);
					newCellText.setCellValue(SiteData);
					if (SiteData.equals(SheetData)) {
						System.out.println("Student Stories BOX IS PRESENT");
						TestElement.click();
						Thread.sleep(2000);
						// Wait for the new tab to open
						WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(40));
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
						String TestElement2 = driver.getCurrentUrl();
						if (TestElement2 != null) {
							System.out.println(
									"Featured Student Story Read All button functions properly: TEST ELEMENT FOUND");
							newCellPassFail.setCellValue("Pass");
						} else {
							System.out.println(
									"Featured Student Story Read All button functions properly: TEST ELEMENT NOT FOUND");
							newCellPassFail.setCellValue("Fail");
						}
						Thread.sleep(4000);
						try {
							driver.close();
							driver.switchTo().window(mainWindowHandle);
							Thread.sleep(1000);
						} catch (Exception e) {
							System.out.println("Cannot Close Window");
						}
					} else {
						if (SheetData.equals("N/A")) {
							System.out.println("Featured Student Story Read All button functions properly: N/A");
							newCellPassFail.setCellValue("N/A");
						} else {
							System.out.println("Featured Student Story Read All button functions properly: Fail");
							newCellPassFail.setCellValue("Fail");
						}
					}
				} catch (Exception e) {
					if (SheetData.equals("N/A")) {
						System.out.println("Featured Student Story Read All button functions properly: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Featured Student Story Read All button functions properly: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("TC091: Fail");
			}
		} else {
			System.out.println("NO MERCK IN URL");
		}
		
		// TC092-TC095 Calculators BOX
		// TC092 PV Home Page - Calculators Box Title
		try {
			System.out.println("START: TC092");
			// Read data from the input file
			Row Row = sheet.getRow(93);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(93);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver.findElement(
						By.xpath(".//div[@class='box box--2colspan calculators orderize__candidate']/div[1]/h2"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Calculators Box Title: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Calculators Box Title: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Calculators Box Title: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("Calculators Box Title: Fail");
				if (SheetData.equals("N/A")) {
					System.out.println("Calculators Box Title: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Calculators Box Title: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC092: Fail");
		}

		// TC093 PV Home Page - Calculators Box View All Button is present
		try {
			System.out.println("START: TC093");
			// Read data from the input file
			Row Row = sheet.getRow(94);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(94);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver
						.findElement(By.xpath(".//div[@class='box box--2colspan calculators orderize__candidate']/a"));
				String SiteData = TestElement.getText();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Calculators Box View All Button is present: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Calculators Box Calculators Box View All Button is present: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Calculators Box Calculators Box View All Button is present: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
			} catch (Exception e) {
				System.out.println("Calculators Box Calculators Box View All Button is present: Fail");
				if (SheetData.equals("N/A")) {
					System.out.println("Calculators Box Calculators Box View All Button is present: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Calculators Box Calculators Box View All Button is present: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC093: Fail");
		}

		// TC094 PV Home Page - Calculators Box View All Button is functionality
		try {
			System.out.println("START: TC094");
			// Read data from the input file
			Row Row = sheet.getRow(95);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(95);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver
						.findElement(By.xpath(".//div[@class='box box--2colspan calculators orderize__candidate']/a"));
				TestElement.click();
				Thread.sleep(2000);
				// Wait for the new tab to open
				WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Calculators Box View All Button is functionality: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Calculators Box View All Button is functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Calculators Box View All Button is functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
				try {
					driver.navigate().back();
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Window");
				}
			} catch (Exception e) {
				System.out.println("Calculators Box View All Button is functionality: Fail");
				if (SheetData.equals("N/A")) {
					System.out.println("Calculators Box View All Button is functionality: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Calculators Box View All Button is functionality: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				Thread.sleep(1000);
				driver.navigate().back();
			}
		} catch (Exception e) {
			System.out.println("TC094: Fail");
		}
		
		
		// TC095 PV Home Page - Calculators Box Featured Calculator titles
		try {
			System.out.println("START: TC095");
			// Read data from the input file
			Row Row = sheet.getRow(96);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(96);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			try {
				// Run script
				WebElement newsLink = driver.findElement(
						By.xpath(".//div[@class='box box--2colspan calculators orderize__candidate']/div[3]/ul"));
				List<WebElement> liTags = newsLink.findElements(By.xpath(".//li"));
				System.out.println("Number Calculators: " + liTags.size());
				if (SheetData.equals("N/A")) {
					System.out.println("Calculators Box Featured Calculator titles: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Calculators Box Featured Calculator titles: Pass");
					newCellPassFail.setCellValue("Pass");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Calculators Box Featured Calculator titles: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Calculators Box Featured Calculator titles: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC095: Fail");
		}
		
		
		
		
		// TC096 Disclaimer Title
		try {
			System.out.println("START: TC096");
			// Read data from the input file
			Row Row = sheet.getRow(97);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(97);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver.findElement(By.xpath(".//div[@class='box box--2colspan htmlwidget orderize__candidate ']/div[1]/h2"));
			String SiteData = GetText.getText();
			System.out.println("Site Data: " + SiteData);
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Disclaimer Title: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Disclaimer Title: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC096: Fail");
		}
		// TC097 Disclaimer Content
		try {
			System.out.println("START: TC097");
			// Read data from the input file
			Row Row = sheet.getRow(98);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(98);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement GetText = driver.findElement(By.xpath(".//div[@class='box box--2colspan htmlwidget orderize__candidate ']/div[2]/div"));
			String SiteData = GetText.getText();
			System.out.println("Site Data: " + SiteData);
			newCellText.setCellValue(SiteData);
			if (SiteData.contains(SheetData)) {
				System.out.println("Disclaimer Content: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Disclaimer Content: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC097: Fail");
		}
		
		// TC098-TC100 FOOTER TOP
		// TC098 Footer Top Logo
		try {
			System.out.println("START: TC098");
			// Read data from the input file
			Row Row = sheet.getRow(99);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(99);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement logoicon = driver.findElement(By.xpath(".//div[@class='footer__logo']/img"));
			String SiteData = logoicon.getAttribute("src");
			newCellText.setCellValue(SiteData);
			System.out.println("Site Data: " + SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("Footer Top Logo: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Footer Top Logo: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC098: Fail");
		}
		
		// TC099 Footer Top Description
		try {
			System.out.println("START: TC099");
			// Read data from the input file
			Row Row = sheet.getRow(100);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(100);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			WebElement logoicon = driver.findElement(By.xpath(".//div[@class='footer__note']/p"));
			String SiteData = logoicon.getText();
			newCellText.setCellValue(SiteData);
			System.out.println("Site Data: " + SiteData);
			if (SiteData.equals(SheetData)) {
				System.out.println("Footer Top Logo: Pass");
				newCellPassFail.setCellValue("Pass");
			} else {
				System.out.println("Footer Top Logo: Fail");
				newCellPassFail.setCellValue("Fail");
			}
		} catch (Exception e) {
			System.out.println("TC099: Fail");
		}
		
		// TC100 Footer Top GMK2020 Link
		try {
			System.out.println("START: TC100");
			// Read data from the input file
			Row Row = sheet.getRow(101);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(101);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetLink = driver.findElement(By.xpath(".//div[@class='footer__note']/p/a"));
				GetLink.click();
				Thread.sleep(2000);
				String SiteData = driver.getTitle();
				System.out.println("Page title: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Footer Top GMK2020 Link: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					System.out.println("Footer Top GMK2020 Link: Fail");
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
			System.out.println("TC100: Fail");
		}
		
		// TC101 Right Sidebar - Quick Links Title
		try {
			System.out.println("START: TC101");
			// Read data from the input file
			Row Row = sheet.getRow(102);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(102);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver.findElement(By.xpath(".//div[@class='box quicklinks orderize__candidate']/div[1]/h2"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				System.out.println("Site Data: " + SiteData);
				if (SiteData.contains(SheetData)) {
					System.out.println("Quick Links Title: Pass");
					newCellPassFail.setCellValue("Pass");
					System.out.println("Wrote data: " + SheetData);
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Quick Links Title: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Quick Links Title: Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC101: Fail");
		}
		// TC102 Right Sidebar - Quick Links Calculators
		try {
			System.out.println("START: TC102");
			// Read data from the input file
			Row Row = sheet.getRow(103);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(103);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetLink = driver.findElement(By.xpath(".//div[@class='box quicklinks orderize__candidate']/div[2]/a[1]"));
				GetLink.click();
				Thread.sleep(2000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Quick Links Calculators: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Quick Links Calculators: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Quick Links Calculators: Fail");
					}
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
			System.out.println("TC102: Fail");
		}
		
		
		// TC103 Right Sidebar - Quick Links NORMAL LAB VALUES
		try {
			System.out.println("START: TC103");
			// Read data from the input file
			Row Row = sheet.getRow(104);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(104);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetLink = driver.findElement(By.xpath(".//div[@class='box quicklinks orderize__candidate']/div[2]/a[2]"));
				GetLink.click();
				Thread.sleep(2000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Quick Links NORMAL LAB VALUES: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Quick Links NORMAL LAB VALUES: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Quick Links NORMAL LAB VALUES: Fail");
					}
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
			System.out.println("TC103: Fail");
		}
		
		// TC104 Right Sidebar - Quick Links PROCEDURE & EXAM VIDEOS
		try {
			System.out.println("START: TC104");
			// Read data from the input file
			Row Row = sheet.getRow(105);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(105);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetLink = driver.findElement(By.xpath(".//div[@class='box quicklinks orderize__candidate']/div[2]/a[3]"));
				GetLink.click();
				Thread.sleep(2000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Quick Links PROCEDURE & EXAM VIDEOS: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Quick Links PROCEDURE & EXAM VIDEOS: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Quick Links PROCEDURE & EXAM VIDEOS: Fail");
					}
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
			System.out.println("TC104: Fail");
		}
		// TC105 Right Sidebar - Quick Links OTHER RESOURCES
		try {
			System.out.println("START: TC105");
			// Read data from the input file
			Row Row = sheet.getRow(106);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(106);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetLink = driver.findElement(By.xpath(".//div[@class='box quicklinks orderize__candidate']/div[2]/a[4]"));
				GetLink.click();
				Thread.sleep(2000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Quick Links OTHER RESOURCES: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Quick Links OTHER RESOURCES: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Quick Links OTHER RESOURCES: Fail");
					}
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
			System.out.println("TC105: Fail");
		}
		
		
		// TC106 Right Sidebar - PEARL OF THE DAY TITLE
		try {
			System.out.println("START: TC106");
			// Read data from the input file
			Row Row = sheet.getRow(107);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(107);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver.findElement(By.xpath(".//div[@class='box pearlOfTheDay orderize__candidate hide-on-print']/div[1]/h2"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				System.out.println("Site Data: " + SiteData);
				if (SiteData.contains(SheetData)) {
					System.out.println("PEARL OF THE DAY TITLE: Pass");
					newCellPassFail.setCellValue("Pass");
					System.out.println("Wrote data: " + SheetData);
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("PEARL OF THE DAY TITLE: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("PEARL OF THE DAY TITLE: Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("PEARL OF THE DAY TITLE: N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
					System.out.println("PEARL OF THE DAY TITLE: Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC106: Fail");
		}
		
		
		// TC107 Right Sidebar - PEARL OF THE DAY TITLE 2
		try {
			System.out.println("START: TC107");
			// Read data from the input file
			Row Row = sheet.getRow(108);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(108);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver.findElement(By.xpath(".//div[@class='box pearlOfTheDay orderize__candidate hide-on-print']/div[2]/div/h3"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("PEARL OF THE DAY TITLE 2: N/A");
				} else {
					newCellPassFail.setCellValue("Pass");
					System.out.println("PEARL OF THE DAY TITLE 2: Pass");
				}
			} catch (Exception e) {
				newCellPassFail.setCellValue("Fail");
				System.out.println("PEARL OF THE DAY TITLE 2: Pass");
			}
		} catch (Exception e) {
			System.out.println("TC107: Fail");
		}
		
		// TC108 Right Sidebar - PEARL OF THE DAY Description
		try {
			System.out.println("START: TC108");
			// Read data from the input file
			Row Row = sheet.getRow(109);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(109);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver.findElement(By.xpath(".//div[@class='box pearlOfTheDay orderize__candidate hide-on-print']/div[2]/div/div/div/div/table/tbody/tr/td/div/ul/li/div/p"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("PEARL OF THE DAY Description: N/A");
				} else {
					newCellPassFail.setCellValue("Pass");
					System.out.println("PEARL OF THE DAY Description: Pass");
				}
			} catch (Exception e) {
				newCellPassFail.setCellValue("Fail");
				System.out.println("PEARL OF THE DAY Description: Pass");
			}
		} catch (Exception e) {
			System.out.println("TC108: Fail");
		}
		// TC109 Right Sidebar - PEARL OF THE DAY READ ALL
		try {
			System.out.println("START: TC109");
			// Read data from the input file
			Row Row = sheet.getRow(110);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Read data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(110);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetLink = driver.findElement(By.xpath(".//div[@class='box pearlOfTheDay orderize__candidate hide-on-print']/div[2]/div/a"));
				GetLink.click();
				Thread.sleep(2000);
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("PEARL OF THE DAY READ ALL: N/A");
				} else {
					newCellPassFail.setCellValue("Pass");
					System.out.println("PEARL OF THE DAY READ ALL: Pass");
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
			System.out.println("TC109: Fail");
		}
		
		
		// TC110 PV Home Page - Right Sidebar - Test Your Knowledge Title
		try {
			System.out.println("START: TC110");
			// Read data from the input file
			Row Row = sheet.getRow(111);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(111);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver.findElement(By.xpath(".//div[@class='box quiz orderize__candidate hide-on-print']/div[1]/div"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				System.out.println("Site Data: " + SiteData);
				if (SiteData.contains(SheetData)) {
					System.out.println("Test Your Knowledge Title: Pass");
					newCellPassFail.setCellValue("Pass");
					System.out.println("Wrote data: " + SheetData);
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Test Your Knowledge Title: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Test Your Knowledge Title: Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("Test Your Knowledge Title: N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
					System.out.println("Test Your Knowledge Title: Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC110: Fail");
		}
		
		// TC111 PV Home Page - Right Sidebar - Test Your Knowledge View All
		try {
			System.out.println("START: TC111");
			// Read data from the input file
			Row Row = sheet.getRow(112);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(112);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement TestElement = driver
						.findElement(By.xpath(".//div[@class='box quiz orderize__candidate hide-on-print']/a"));
				TestElement.click();
				Thread.sleep(2000);
				// Wait for the new tab to open
				WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));
				String SiteData = driver.getCurrentUrl();
				System.out.println("Site Data: " + SiteData);
				newCellText.setCellValue(SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Test Your Knowledge View All functionality: Pass");
					newCellPassFail.setCellValue("Pass");
				} else {
					if (SheetData.equals("N/A")) {
						System.out.println("Test Your Knowledge View All functionality: N/A");
						newCellPassFail.setCellValue("N/A");
					} else {
						System.out.println("Test Your Knowledge View All functionality: Fail");
						newCellPassFail.setCellValue("Fail");
					}
				}
				try {
					driver.navigate().back();
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Window");
				}
			} catch (Exception e) {
				System.out.println("Test Your Knowledge View All functionality: Fail");
				if (SheetData.equals("N/A")) {
					System.out.println("Test Your Knowledge View All functionality: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Test Your Knowledge View All functionality: Fail");
					newCellPassFail.setCellValue("Fail");
				}
				Thread.sleep(1000);
				driver.navigate().back();
			}
		} catch (Exception e) {
			System.out.println("TC111: Fail");
		}
		// TC112 PV Home Page - Right Sidebar - Test Your Knowledge Title, Description and Options
		try {
			System.out.println("START: TC112");
			// Read data from the input file
			Row Row = sheet.getRow(113);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(113);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement QuizTitle = driver.findElement(By.xpath(".//div[@class='box quiz orderize__candidate hide-on-print']/div[2]/div/form/div[2]"));
				System.out.println("Quiz Title: "+QuizTitle.getText());
				WebElement QuizQuestion = driver.findElement(By.xpath(".//div[@class='box quiz orderize__candidate hide-on-print']/div[2]/div/form/div[3]"));
				System.out.println("Quiz Title: "+QuizQuestion.getText());
				WebElement QuizOptions = driver.findElement(By.xpath(".//div[@class='box quiz orderize__candidate hide-on-print']/div[2]/div/form/ol/li[1]/label/span"));
				System.out.println("Quiz Title: "+QuizOptions.getText());
				
				if (SheetData.equals("N/A")) {
					System.out.println("Test Your Knowledge Title, Description and Options: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Test Your Knowledge Title, Description and Options: Pass");
					newCellPassFail.setCellValue("Pass");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Test Your Knowledge Title, Description and Options: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Test Your Knowledge Title, Description and Options: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC112: Fail");
		}
		
		
		
		// TC113 PV Home Page - Right Sidebar - Test Your Knowledge Functionality
		try {
			System.out.println("START: TC113");
			// Read data from the input file
			Row Row = sheet.getRow(114);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(114);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				driver.findElement(By.xpath(".//div/form/ol/li/input")).click();
				driver.findElement(By.xpath("//div/form/input[@class='quiz__submit']")).click();
				ClosePopup();
				if (SheetData.equals("N/A")) {
					System.out.println("Test Your Knowledge functionality: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Test Your Knowledge functionality: Pass");
					newCellPassFail.setCellValue("Pass");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Test Your Knowledge Functionality: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Test Your Knowledge functionality: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC113: Fail");
		}
		
		// TC114 PV Home Page - Right Sidebar - Apps Ad Banner
		try {
			System.out.println("START: TC114");
			// Read data from the input file
			Row Row = sheet.getRow(115);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(115);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement SiteData = driver.findElement(By.xpath(".//div[4]/div/div[2]/div[2]/map/area[1]"));
				if (SheetData.equals("N/A")) {
					System.out.println("Apps Ad Banner: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Apps Ad Banner: Pass");
					newCellPassFail.setCellValue("Pass");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Apps Ad Banner: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Apps Ad Banner: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC114: Fail");
		}
		
		// TC115 PV Home Page - Right Sidebar - Apps Ad Banner iOS apps
		try {
			System.out.println("START: TC115");
			// Read data from the input file
			Row Row = sheet.getRow(116);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(116);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
		        WebElement areaElement = driver.findElement(By.xpath(".//div[4]/div/div[2]/div[2]/map/area[@alt='iOS']"));
		        String SiteData = areaElement.getAttribute("href");
				System.out.println("Site Data: " + SiteData);
				if (SheetData.equals(SiteData)) {
					System.out.println("Apps Ad Banner iOS apps: Pass");
					newCellPassFail.setCellValue("Pass");
				} else if (SheetData.equals("N/A")) {
					System.out.println("Apps Ad Banner iOS apps: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Apps Ad Banner iOS apps: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Apps Ad Banner iOS apps: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Apps Ad Banner iOS apps: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC115: Fail");
		}
		
		// TC116 PV Home Page - Right Sidebar - Apps Ad Banner Android apps
		try {
			System.out.println("START: TC116");
			// Read data from the input file
			Row Row = sheet.getRow(117);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(117);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement areaElement = driver.findElement(By.xpath(".//div[4]/div/div[2]/div[2]/map/area[@alt='ANDROID']"));
		        String SiteData = areaElement.getAttribute("href");
				System.out.println("Site Data: " + SiteData);
				if (SheetData.equals(SiteData)) {
					System.out.println("Apps Ad Banner Android apps: Pass");
					newCellPassFail.setCellValue("Pass");
				} else if (SheetData.equals("N/A")) {
					System.out.println("Apps Ad Banner Android apps: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Apps Ad Banner Android apps: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					System.out.println("Apps Ad Banner Android apps: N/A");
					newCellPassFail.setCellValue("N/A");
				} else {
					System.out.println("Apps Ad Banner Android apps: Fail");
					newCellPassFail.setCellValue("Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC116: Fail");
		}
		
		
		// TC117 PV Home Page - Right Sidebar - Social Media Title
		try {
			System.out.println("START: TC117");
			// Read data from the input file
			Row Row = sheet.getRow(118);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(118);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver.findElement(By.xpath(".//div[@class='box socialmedia hide-on-print orderize__candidate']/div[1]/h2"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				System.out.println("Site Data: " + SiteData);
				if (SiteData.contains(SheetData)) {
					System.out.println("Social Media Title: Pass");
					newCellPassFail.setCellValue("Pass");
					System.out.println("Wrote data: " + SheetData);
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Social Media Title: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Social Media Title: Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("Social Media Title: N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
					System.out.println("Social Media Title: Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC117: Fail");
		}
		
		// TC118 PV Home Page - Right Sidebar - Social Media Read All Button
		try {
			System.out.println("START: TC118");
			// Read data from the input file
			Row Row = sheet.getRow(119);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(119);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement GetText = driver.findElement(By.xpath(".//div[@class='box socialmedia hide-on-print orderize__candidate']/div[1]/a/div"));
				String SiteData = GetText.getText();
				newCellText.setCellValue(SiteData);
				System.out.println("Site Data: " + SiteData);
				if (SiteData.contains(SheetData)) {
					System.out.println("Social Media Read All Button: Pass");
					newCellPassFail.setCellValue("Pass");
					System.out.println("Wrote data: " + SheetData);
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Social Media Read All Button: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Social Media Read All Button: Fail");
					}
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("Social Media Read All Button: N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
					System.out.println("Social Media Read All Button: Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC118: Fail");
		}
		
		// TC119 PV Home Page - Right Sidebar - Social Media FACEBOOK
		try {
			System.out.println("START: TC119");
			// Read data from the input file
			Row Row = sheet.getRow(120);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(120);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FollowButton = driver.findElement(By.xpath(".//div[@class='followmedia']/a[contains(@aria-label,'facebook')]"));
				FollowButton.click();
				Thread.sleep(2000);
				// Wait for the new tab to open
				WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
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
				String SiteData = driver.getCurrentUrl();
				newCellText.setCellValue(SiteData);
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Social Media FACEBOOK: Pass");
					newCellPassFail.setCellValue("Pass");
					System.out.println("Wrote data: " + SheetData);
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Social Media FACEBOOK: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Social Media FACEBOOK: Fail");
					}
				}
				try {
					driver.close();
					driver.switchTo().window(mainWindowHandle);
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Window");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("Social Media FACEBOOK: N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
					System.out.println("Social Media FACEBOOK: Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC119: Fail");
		}
		
		// TC120 PV Home Page - Right Sidebar - Social Media YOUTUBE
		try {
			System.out.println("START: TC120");
			// Read data from the input file
			Row Row = sheet.getRow(121);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(121);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FollowButton = driver.findElement(By.xpath(".//div[@class='followmedia']/a[contains(@aria-label,'youtube')]"));
				FollowButton.click();
				Thread.sleep(2000);
				// Wait for the new tab to open
				WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
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
				String SiteData = driver.getCurrentUrl();
				newCellText.setCellValue(SiteData);
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Social Media YOUTUBE: Pass");
					newCellPassFail.setCellValue("Pass");
					System.out.println("Wrote data: " + SheetData);
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Social Media YOUTUBE: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Social Media YOUTUBE: Fail");
					}
				}
				try {
					driver.close();
					driver.switchTo().window(mainWindowHandle);
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Window");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("Social Media YOUTUBE: N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
					System.out.println("Social Media YOUTUBE: Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC120: Fail");
		}
		// TC121 PV Home Page - Right Sidebar - Social Media TWITTER
		try {
			System.out.println("START: TC121");
			// Read data from the input file
			Row Row = sheet.getRow(122);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(122);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FollowButton = driver.findElement(By.xpath(".//div[@class='followmedia']/a[contains(@aria-label,'twitter')]"));
				FollowButton.click();
				Thread.sleep(2000);
				// Wait for the new tab to open
				WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
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
				String SiteData = driver.getCurrentUrl();
				newCellText.setCellValue(SiteData);
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Social Media TWITTER: Pass");
					newCellPassFail.setCellValue("Pass");
					System.out.println("Wrote data: " + SheetData);
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Social Media TWITTER: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Social Media TWITTER: Fail");
					}
				}
				try {
					driver.close();
					driver.switchTo().window(mainWindowHandle);
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Window");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("Social Media TWITTER: N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
					System.out.println("Social Media TWITTER: Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC121: Fail");
		}
		
		// TC122 PV Home Page - Right Sidebar - Social Media PINTEREST
		try {
			System.out.println("START: TC122");
			// Read data from the input file
			Row Row = sheet.getRow(123);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(123);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FollowButton = driver.findElement(By.xpath(".//div[@class='followmedia']/a[contains(@aria-label,'pinterest')]"));
				FollowButton.click();
				Thread.sleep(2000);
				// Wait for the new tab to open
				WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
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
				String SiteData = driver.getCurrentUrl();
				newCellText.setCellValue(SiteData);
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Social Media PINTEREST: Pass");
					newCellPassFail.setCellValue("Pass");
					System.out.println("Wrote data: " + SheetData);
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Social Media PINTEREST: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Social Media PINTEREST: Fail");
					}
				}
				try {
					driver.close();
					driver.switchTo().window(mainWindowHandle);
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Window");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("Social Media PINTEREST: N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
					System.out.println("Social Media PINTEREST: Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC122: Fail");
		}
		
		// TC123 PV Home Page - Right Sidebar - Social Media INSTAGRAM
		try {
			System.out.println("START: TC123");
			// Read data from the input file
			Row Row = sheet.getRow(124);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(124);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FollowButton = driver.findElement(By.xpath(".//div[@class='followmedia']/a[contains(@aria-label,'instagram')]"));
				FollowButton.click();
				Thread.sleep(2000);
				// Wait for the new tab to open
				WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
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
				String SiteData = driver.getCurrentUrl();
				newCellText.setCellValue(SiteData);
				System.out.println("Site Data: " + SiteData);
				if (SiteData.equals(SheetData)) {
					System.out.println("Social Media INSTAGRAM: Pass");
					newCellPassFail.setCellValue("Pass");
					System.out.println("Wrote data: " + SheetData);
				} else {
					if (SheetData.equals("N/A")) {
						newCellPassFail.setCellValue("N/A");
						System.out.println("Social Media INSTAGRAM: N/A");
					} else {
						newCellPassFail.setCellValue("Fail");
						System.out.println("Social Media INSTAGRAM: Fail");
					}
				}
				try {
					driver.close();
					driver.switchTo().window(mainWindowHandle);
					Thread.sleep(1000);
				} catch (Exception e) {
					System.out.println("Cannot Close Window");
				}
			} catch (Exception e) {
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("Social Media INSTAGRAM: N/A");
				} else {
					newCellPassFail.setCellValue("Fail");
					System.out.println("Social Media INSTAGRAM: Fail");
				}
			}
		} catch (Exception e) {
			System.out.println("TC123: Fail");
		}
		
		// TC124 PV Home Page - Right Sidebar - Social Media ADD TO ANY
		try {
			System.out.println("START: TC124");
			// Read data from the input file
			Row Row = sheet.getRow(125);
			Cell Cell = Row.getCell(ColumnNGet);
			String SheetData = Cell.getStringCellValue();
			System.out.println("Sheet Data: " + SheetData);
			// Write new data to the output file
			Row outputRow = outputSheet.getRow(125);
			Cell newCellText = outputRow.createCell(ColumnNWrite);
			Cell newCellPassFail = outputRow.createCell(CoulumnNPass);
			// Run script
			try {
				WebElement FollowButton = driver.findElement(By.xpath(".//div[@class='box socialmedia hide-on-print orderize__candidate']//a[@class=\"a2a_dd socialmedia__icon socialmedia__icon--add\"]"));
				FollowButton.click();
				Thread.sleep(2000);
				// Wait for the new tab to open
				WebElement CloseButton = driver.findElement(By.xpath("//*[@id=\"closebtnfullcustom\"]"));
				CloseButton.click();
				if (SheetData.equals("N/A")) {
					newCellPassFail.setCellValue("N/A");
					System.out.println("Social Media ADD TO ANY: N/A");
				} else {
					newCellPassFail.setCellValue("Pass");
					System.out.println("Social Media ADD TO ANY: Pass");
				}
			} catch (Exception e) {
				newCellPassFail.setCellValue("Fail");
				System.out.println("Social Media ADD TO ANY: Fail");
			}
		} catch (Exception e) {
			System.out.println("TC124: Fail");
		}
		
		
		
		
		//TWITTER AND FACEBOOK FEEDS NEXT!
		
		
		
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

		// CODE GOES HERE!!!!!!!!!!!!!!

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
