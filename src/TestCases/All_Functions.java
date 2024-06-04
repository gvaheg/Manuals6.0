package TestCases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class All_Functions {

	// START DRIVER
	static WebDriver driver;
	static XSSFWorkbook wb;

	
	
	// HEADER Testing
	public void verifyHeader() throws Exception {

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
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		//END CODE GOES HERE!!!!!!!!!!!!!!
		// Save the changes to the output file
		workbook.close();
		file.close();
		outputWorkbook.write(output);
		output.close();
		outputWorkbook.close();
		System.out.println("Wrote Data Successfully!");

	}

	
	
	// FOOTER Testing
	public void verifyFooter() throws Exception {

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
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		//END CODE GOES HERE!!!!!!!!!!!!!!
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
