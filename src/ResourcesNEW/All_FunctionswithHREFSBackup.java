package ResourcesNEW;


import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.platform.commons.function.Try;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.UnreachableBrowserException;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

public class All_FunctionswithHREFSBackup {

	private static final boolean Figures = false;
    // START DRIVER
    static WebDriver wd;
    static XSSFWorkbook wb;
    private Map<String, Integer> hreflangColumnMap = new HashMap<>();

    // Setter for WebDriver
    public static void setWebDriver(WebDriver driver) {
        wd = driver;
        
    }
	

	// VERIFY FIGURES
    public void verifyFigures() throws Exception {
        ManageWindows();
        WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
        WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
        WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(10));
        WebDriverWait wait5 = new WebDriverWait(wd, Duration.ofSeconds(5));
        String currentURL = wd.getCurrentUrl();
        wd.get(currentURL);
        confirmPage(wait20, wait50);
        CloseCookies();
        try {
            int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
            System.out.println("We have Sections: " + sectionCount);
            int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
            System.out.println("Total Number: " + TotalCount);
            XSSFSheet sheet = wb.createSheet();
            Row rowHeading4 = sheet.createRow(4);
            int HeadN = 0;
            rowHeading4.createCell(HeadN++).setCellValue("SECTION");
            rowHeading4.createCell(HeadN++).setCellValue("ID");
            rowHeading4.createCell(HeadN++).setCellValue("TITLE");
            rowHeading4.createCell(HeadN++).setCellValue("CREDITS");
            rowHeading4.createCell(HeadN++).setCellValue("DESCRIPTION");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
            rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
            rowHeading4.createCell(HeadN++).setCellValue("FILE URL");
            rowHeading4.createCell(HeadN++).setCellValue("FILE PATH");
            rowHeading4.createCell(HeadN++).setCellValue("IMAGE");
            excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);
            CloseCookies();
            int rowNum = 5;

            for (int s = 2; s < 36; s++) {
                // SECTIONS>
                String sectionText = null;
                while (true) {
                    try {
                        if (s > 36) {
                            break; // Exit the loop if s exceeds 34
                        }
                        WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                        sectionText = section.getAttribute("innerText"); // Store section text
                        break; // Exit the loop if the section is found
                    } catch (Exception e) {
                        System.out.println("Can't Find Section for s = " + s);
                        s++; // Increment s if the section is not found
                    }
                }
                // <END SECTIONS
                int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                System.out.println("Rows in this section: " + rowCount);
                // Main For Loop
                for (int i = 1; i < rowCount + 1; i++) {
                    try {
                        Row rowN = sheet.createRow(rowNum++);
                        // Section Text to Excel
                        rowN.createCell(0).setCellValue(sectionText);
                        System.out.println("Section: " + sectionText);
                        rowN.createCell(1).setCellValue(i);
                        // Run Main Codes
                        openResource(rowN, i, s);
                        ShowDetails();
                        getCredits(rowN, i);
                        getFileName(rowN, sheet);
                        getLocation(rowN, sheet, wait10, wait20, wait50, i, rowN, HeadN, i);
                        closePopup();
                    } catch (Exception e) {
                        System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
                        Thread.sleep(3000);
                        ((JavascriptExecutor) wd).executeScript("window.open()");
                        Thread.sleep(2000);
                        ArrayList<String> tabs1 = new ArrayList<String>(wd.getWindowHandles());
                        Thread.sleep(2000);
                        wd.close();
                        wd.switchTo().window(tabs1.get(1));
                        Thread.sleep(2000);
                        wd.get(currentURL);
                        Thread.sleep(5000);
                        CloseCookies();
                        return;
                    }
                    try {
                        FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Figures.xlsx");
                        wb.write(fout);
                        fout.close();
                        System.out.println("Written Successfully!");
                    } catch (Exception e) {
                        System.out.println("Cannot SAVE File!");
                    }
                }
            }

        } catch (Exception e) {
            System.out.println("ERROR! Web page is not responding. Reopening the page... ");
            ErrorMain(currentURL, e);
        }
    }
 

	// VERIFY IMAGES
    public void verifyImages() throws Exception {
        ManageWindows();
        WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
        WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
        WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(10));
        WebDriverWait wait5 = new WebDriverWait(wd, Duration.ofSeconds(5));
        String currentURL = wd.getCurrentUrl();
        wd.get(currentURL);
        confirmPage(wait20, wait50);
        CloseCookies();
        try {
            int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
            System.out.println("We have Sections: " + sectionCount);
            int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
            System.out.println("Total Number: " + TotalCount);
            XSSFSheet sheet = wb.createSheet();
            Row rowHeading4 = sheet.createRow(4);
            int HeadN = 0;
            rowHeading4.createCell(HeadN++).setCellValue("SECTION");
            rowHeading4.createCell(HeadN++).setCellValue("ID");
            rowHeading4.createCell(HeadN++).setCellValue("TITLE");
            rowHeading4.createCell(HeadN++).setCellValue("CREDITS");
            rowHeading4.createCell(HeadN++).setCellValue("DESCRIPTION");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
            rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
            rowHeading4.createCell(HeadN++).setCellValue("FILE URL");
            rowHeading4.createCell(HeadN++).setCellValue("FILE PATH");
            rowHeading4.createCell(HeadN++).setCellValue("IMAGE");
            excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);
            CloseCookies();
            int rowNum = 5;

            for (int s = 2; s < 36; s++) {
                // SECTIONS>
                String sectionText = null;
                while (true) {
                    try {
                        if (s > 36) {
                            break; // Exit the loop if s exceeds 34
                        }
                        WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                        sectionText = section.getAttribute("innerText"); // Store section text
                        break; // Exit the loop if the section is found
                    } catch (Exception e) {
                        System.out.println("Can't Find Section for s = " + s);
                        s++; // Increment s if the section is not found
                    }
                }
                // <END SECTIONS
                int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                System.out.println("Rows in this section: " + rowCount);
                // Main For Loop
                for (int i = 1; i < rowCount + 1; i++) {
                    try {
                        Row rowN = sheet.createRow(rowNum++);
                        // Section Text to Excel
                        rowN.createCell(0).setCellValue(sectionText);
                        System.out.println("Section: " + sectionText);
                        rowN.createCell(1).setCellValue(i);
                        // Run Main Codes
                        openResource(rowN, i, s);
                        ShowDetails();
                        getCredits(rowN, i);
                        getFileName(rowN, sheet);
                        getLocation(rowN, sheet, wait10, wait20, wait50, i, rowN, HeadN, i);
                        closePopup();
                    } catch (Exception e) {
                        System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
                        Thread.sleep(3000);
                        ((JavascriptExecutor) wd).executeScript("window.open()");
                        Thread.sleep(2000);
                        ArrayList<String> tabs1 = new ArrayList<String>(wd.getWindowHandles());
                        Thread.sleep(2000);
                        wd.close();
                        wd.switchTo().window(tabs1.get(1));
                        Thread.sleep(2000);
                        wd.get(currentURL);
                        Thread.sleep(5000);
                        CloseCookies();
                        return;
                    }
                    try {
                        FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Images.xlsx");
                        wb.write(fout);
                        fout.close();
                        System.out.println("Written Successfully!");
                    } catch (Exception e) {
                        System.out.println("Cannot SAVE File!");
                    }
                }
            }

        } catch (Exception e) {
            System.out.println("ERROR! Web page is not responding. Reopening the page... ");
            ErrorMain(currentURL, e);
        }
    }
	

	// VERIFY SOUNDS
    public void verifySounds() throws Exception {
        ManageWindows();
        WebDriverWait wait50 = new WebDriverWait(wd, Duration.ofSeconds(50));
        WebDriverWait wait20 = new WebDriverWait(wd, Duration.ofSeconds(20));
        WebDriverWait wait10 = new WebDriverWait(wd, Duration.ofSeconds(10));
        WebDriverWait wait5 = new WebDriverWait(wd, Duration.ofSeconds(5));
        String currentURL = wd.getCurrentUrl();
        wd.get(currentURL);
        confirmPage(wait20, wait50);
        CloseCookies();
        try {
            int sectionCount = wd.findElements(By.xpath(".//div/div/div/div/h2")).size();
            System.out.println("We have Sections: " + sectionCount);
            int TotalCount = wd.findElements(By.xpath(".//div/div/div/div/div[1]/div/div/div[2]/ul/li")).size();
            System.out.println("Total Number: " + TotalCount);
            XSSFSheet sheet = wb.createSheet();
            Row rowHeading4 = sheet.createRow(4);
            int HeadN = 0;
            rowHeading4.createCell(HeadN++).setCellValue("SECTION");
            rowHeading4.createCell(HeadN++).setCellValue("ID");
            rowHeading4.createCell(HeadN++).setCellValue("TITLE");
            rowHeading4.createCell(HeadN++).setCellValue("CREDITS");
            rowHeading4.createCell(HeadN++).setCellValue("DESCRIPTION");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION TITLE");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION (PASS/FAIL)");
            rowHeading4.createCell(HeadN++).setCellValue("LOCATION URL");
            rowHeading4.createCell(HeadN++).setCellValue("LOC. URL RESPONSE");
            //rowHeading4.createCell(HeadN++).setCellValue("FILE URL");
            //rowHeading4.createCell(HeadN++).setCellValue("FILE PATH");
            //rowHeading4.createCell(HeadN++).setCellValue("IMAGE");
            excelRows(sheet, sectionCount, TotalCount, HeadN, rowHeading4);
            CloseCookies();
            int rowNum = 5;

            for (int s = 2; s < 36; s++) {
                // SECTIONS>
                String sectionText = null;
                while (true) {
                    try {
                        if (s > 36) {
                            break; // Exit the loop if s exceeds 34
                        }
                        WebElement section = wd.findElement(By.xpath(".//main/div[1]//div[" + s + "]//div/div/h2"));
                        sectionText = section.getAttribute("innerText"); // Store section text
                        break; // Exit the loop if the section is found
                    } catch (Exception e) {
                        System.out.println("Can't Find Section for s = " + s);
                        s++; // Increment s if the section is not found
                    }
                }
                // <END SECTIONS
                int rowCount = wd.findElements(By.xpath(".//main/div[1]//div[" + s + "]//div/div[2]/ul/li")).size();
                System.out.println("Rows in this section: " + rowCount);
                // Main For Loop
                for (int i = 1; i < rowCount + 1; i++) {
                    try {
                        Row rowN = sheet.createRow(rowNum++);
                        // Section Text to Excel
                        rowN.createCell(0).setCellValue(sectionText);
                        System.out.println("Section: " + sectionText);
                        rowN.createCell(1).setCellValue(i);
                        // Run Main Codes
                        openResource(rowN, i, s);
                        ShowDetails();
                        getCredits(rowN, i);
                        //getFileName(rowN, sheet);
                        getLocation(rowN, sheet, wait10, wait20, wait50, i, rowN, HeadN, i);
                        closePopup();
                    } catch (Exception e) {
                        System.out.println("ERROR! Resource page is not responding. Reopening the page... ");
                        Thread.sleep(3000);
                        ((JavascriptExecutor) wd).executeScript("window.open()");
                        Thread.sleep(2000);
                        ArrayList<String> tabs1 = new ArrayList<String>(wd.getWindowHandles());
                        Thread.sleep(2000);
                        wd.close();
                        wd.switchTo().window(tabs1.get(1));
                        Thread.sleep(2000);
                        wd.get(currentURL);
                        Thread.sleep(5000);
                        CloseCookies();
                        return;
                    }
                    try {
                        FileOutputStream fout = new FileOutputStream("C:\\TestResults\\Sounds.xlsx");
                        wb.write(fout);
                        fout.close();
                        System.out.println("Written Successfully!");
                    } catch (Exception e) {
                        System.out.println("Cannot SAVE File!");
                    }
                }
            }

        } catch (Exception e) {
            System.out.println("ERROR! Web page is not responding. Reopening the page... ");
            ErrorMain(currentURL, e);
        }
    }
	//////////////////////////REUSABLE METHODS
	
	
	// Manage Windows
	public void ManageWindows(){
		//wd.manage().window().maximize();
		wd.manage().window().setSize(new Dimension(400, 820));
		wd.manage().window().setPosition(new Point(780, 0));
		//wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(1));
		Actions action = new Actions(wd);
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
	// Close Cookies
	public void CloseCookies() throws InterruptedException {
		try {
			WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
			AcceptCookies.click();
		} catch (Exception e) {
			System.out.println("Can't Close Cookies");
		}
		Thread.sleep(1000);
	}
	
	// Confirm page opens
		public void confirmPage(WebDriverWait wait50,WebDriverWait wait20) {
			try {
				wait20.until(
						ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
			} catch (Exception e) {
				System.out.println("Refreshing the page!");
				wd.navigate().refresh();
				wait20.until(
						ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[1]/div/div/div[2]/ul/li[1]/a")));
			}
		}
	
		public void excelRows(XSSFSheet sheet, int sectionCount, int TotalCount, int HeadN, Row rowHeading4) {
		    try {
		        // Get Date
		        DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
		        Date date = new Date();
		        String date1 = dateFormat.format(date);

		        // Main Excel Rows for all
		        Row rowHeading0 = sheet.createRow(0);
		        Row rowHeading1 = sheet.createRow(1);
		        Row rowHeading2 = sheet.createRow(2);
		        Row rowHeading3 = sheet.createRow(3);

		        rowHeading0.createCell(0).setCellValue("TOTAL NUMBER: " + TotalCount);
		        rowHeading0.createCell(4).setCellValue("TOTAL SECTIONS: " + sectionCount);
		        rowHeading1.createCell(0).setCellValue("PAGE TITLE: " + wd.getTitle());
		        rowHeading2.createCell(0).setCellValue("START DATE/TIME: " + date1);
		        rowHeading3.createCell(0).setCellValue("PAGE URL: " + wd.getCurrentUrl());

		        // Define HREF column headers in the correct order
		        String[] hreflangLabels = {
		            "EN HREF", "DE HREF", "ES HREF", "FR HREF", "IT HREF",
		            "JA HREF", "PT HREF", "RU HREF", "VI HREF", "ZH HREF",
		            "AR HREF", "UK HREF"
		        };

		        // Map HREF columns dynamically
		        for (String hreflang : hreflangLabels) {
		            rowHeading4.createCell(HeadN).setCellValue(hreflang);
		            hreflangColumnMap.put(hreflang.split(" ")[0].toLowerCase(), HeadN++);
		        }
		    } catch (Exception e) {
		        System.out.println("Cannot Create Excel Rows: " + e.getMessage());
		    }
		}


		
	
	// Open Each Resource Item
	public void openResource(Row rowN, int i, int s) throws InterruptedException {
	    int maxTestsPerSession = 30;
	    if (i % maxTestsPerSession == 0) {
	        String currentURL = wd.getCurrentUrl();
	     // Quit the current driver
	        wd.quit();

	        // Initialize a new Firefox driver in headless mode
	        FirefoxOptions options = new FirefoxOptions();
	        if (i >= maxTestsPerSession) {
	            System.out.println("Switching to headless mode...");
	            options.setHeadless(true);
	        } else {
	            System.out.println("Running in normal mode...");
	            options.setHeadless(false);
	        }

	        // Restart the driver
	        wd = new FirefoxDriver(options);
	        Thread.sleep(1000);

	        // Navigate to the last URL
	        wd.get(currentURL);
	        Thread.sleep(2000);

	        // Perform any necessary actions to reset or prepare the session
	        ManageWindows();
	        CloseCookies();
	    }

	    // Main Title
	    try {
	        WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(10));
	        WebElement titleElement = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
	            ".//main/div[1]//div[" + s + "]//div/div[2]/ul/li[" + i + "]/a"
	        )));
	        String titleText = titleElement.getAttribute("innerText");
	        System.out.println("Title " + i + ": " + titleText);

	        // Write ID and Title
	        rowN.createCell(1).setCellValue(i);
	        rowN.createCell(2).setCellValue(titleText);

	        // Scroll to the resource item and click it
	        JavascriptExecutor js = (JavascriptExecutor) wd;
	        js.executeScript("arguments[0].scrollIntoView({block: 'center'});", titleElement);
	        Thread.sleep(500); // Small pause to ensure scrolling completes

	        int retries = 3; // Retry mechanism
	        while (retries > 0) {
	            try {
	                titleElement.click();
	                System.out.println("Clicked on resource item " + i);
	                break; // Exit loop if successful
	            } catch (Exception e) {
	                System.out.println("Failed to click on resource item " + i + ", retrying...");
	                retries--;
	                Thread.sleep(1000);
	            }
	        }
	        if (retries == 0) {
	            throw new Exception("Failed to click on resource item " + i + " after retries");
	        }

	    } catch (Exception e) {
	        System.out.println("Failed to click on resource item " + i);
	        rowN.createCell(1).setCellValue(i);
	        rowN.createCell(2).setCellValue("CANNOT GET TITLE");
	        wd.navigate().refresh();
	    }
	}

	
	
	// Get Credits and Descriptions
	public void getCredits(Row rowN, int i) {

		// Get & write Description
		try {
			WebElement description = wd.findElement(By.xpath(
					"//div[@class='SliderMobile_description__7xhlL']/div/div/p"));
			System.out.println("Description: " + description.getAttribute("innerText"));
			rowN.createCell(4).setCellValue(description.getAttribute("innerText"));
		} catch (Exception e) {
			System.out.println("Description: CANT FIND");
			rowN.createCell(4).setCellValue("");
		}
		// Get & write Credits
		try {
			WebElement credits = wd
					.findElement(By.xpath("//div[@class='SliderMobile_credits___BXrj']/div/div/p"));
			System.out.println("Credits: " + credits.getAttribute("innerText"));
			rowN.createCell(3).setCellValue(credits.getAttribute("innerText"));
		} catch (Exception e) {
			System.out.println("");
			rowN.createCell(3).setCellValue("");
		}

	}
	

	public void getFileName(Row rowN, XSSFSheet sheet) {
	    try {
	        // Increase timeout to handle slower image loading
	        WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(20));

	        // Locate the image element
	        WebElement FileName = wait.until(ExpectedConditions.visibilityOfElementLocated(
	            By.cssSelector(".active.carousel-item .PopupContent_multiMediaImageContainer__wunjw > div.PopupContent_imageCarousel__uy4Mq > img")
	        ));

	        // Extract the URL from the src attribute
	        String srcAttribute = FileName.getAttribute("src");
	        URL FileURL = new URL(srcAttribute);

	        System.out.println("Src attribute is: " + FileURL);
	        System.out.println("Path = " + FileURL.getPath());

	        // Write image details to the Excel sheet
	        rowN.createCell(9).setCellValue(srcAttribute);
	        rowN.createCell(10).setCellValue(FileURL.getPath());

	        // VERIFY IMAGE EXISTS
	        if (verifyImageURLWithRetries(FileURL)) {
	            System.out.println("Image: FOUND");
	            rowN.createCell(11).setCellValue("FOUND");
	        } else {
	            System.out.println("Image: NOT FOUND");
	            rowN.createCell(11).setCellValue("NOT FOUND");
	            highlightCellRed(sheet, rowN.getCell(11));
	        }
	    } catch (Exception e) {
	        System.out.println("CANNOT GET FILE NAME: " + e.getMessage());
	        rowN.createCell(9).setCellValue("No File PATH");
	        rowN.createCell(10).setCellValue("No File URL");
	        rowN.createCell(11).setCellValue("NOT FOUND");
	        highlightCellRed(sheet, rowN.getCell(11));
	    }
	}

	private boolean verifyImageURLWithRetries(URL fileURL) {
	    int retries = 3; // Number of retry attempts
	    while (retries > 0) {
	        try {
	            // Introduce a small delay between retries
	            Thread.sleep(500);

	            // Open a connection to the URL
	            HttpURLConnection connection = (HttpURLConnection) fileURL.openConnection();

	            // Set headers to mimic a browser
	            connection.setRequestMethod("GET");
	            connection.setRequestProperty("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
	            connection.setRequestProperty("Accept", "image/webp,*/*");
	            connection.setConnectTimeout(5000);
	            connection.setReadTimeout(5000);
	            connection.setInstanceFollowRedirects(true); // Follow redirects

	            // Get the response code
	            int responseCode = connection.getResponseCode();

	            if (responseCode >= 200 && responseCode < 300) {
	                System.out.println("Image exists with response code: " + responseCode);
	                return true; // Image exists
	            } else {
	                System.out.println("Image check failed with response code: " + responseCode);
	            }
	        } catch (IOException e) {
	            System.out.println("Error verifying URL: " + e.getMessage());
	        } catch (InterruptedException ie) {
	            Thread.currentThread().interrupt();
	            System.out.println("Thread interrupted during retry.");
	        }
	        retries--;
	    }
	    return false; // Image verification failed after retries
	}


	
	
	

	// Highlight a cell in red
	private void highlightCellRed(XSSFSheet sheet, Cell cell) {
	    CellStyle style = sheet.getWorkbook().createCellStyle();
	    style.setFillForegroundColor(IndexedColors.RED.getIndex());
	    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    cell.setCellStyle(style);
	}

	// Highlight a cell in yellow
	private void highlightCellYellow(XSSFSheet sheet, Cell cell) {
	    CellStyle style = sheet.getWorkbook().createCellStyle();
	    style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
	    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    cell.setCellStyle(style);
	}


	// Helper method to verify the image URL
	private boolean verifyImageURL(URL fileURL) {
	    try {
	        HttpURLConnection connection = (HttpURLConnection) fileURL.openConnection();
	        connection.setRequestMethod("HEAD");
	        connection.setConnectTimeout(5000); // Timeout for faster verification
	        connection.setReadTimeout(5000);
	        int responseCode = connection.getResponseCode();

	        // Check if the response code indicates the image exists
	        if (responseCode >= 200 && responseCode < 300) {
	            return true;
	        }
	    } catch (Exception e) {
	        System.out.println("Error verifying image URL: " + e.getMessage());
	    }
	    return false;
	}

	
	// ShowDetails
		public void ShowDetails() throws InterruptedException {
			try {
				WebElement showDetails = null;
				try {
					showDetails = wd.findElement(By.xpath("//span[@data-testid='showDetails']"));
				} catch (Exception e) {
					Thread.sleep(1000);
					showDetails = wd.findElement(By.xpath("//*[@id='iframepopupundefined']/div[2]/div[2]/div[1]/div[1]/span[2]"));
				}
				if (showDetails != null) {
				showDetails.click();
				}
			} catch (Exception ex) {
				Thread.sleep(2000);
				System.out.println("Can't CLICK ON SHOW DETAILS");
			}
		}
	//HANDLING ERRORS
	public void ErrorMain(String currentURL, Exception e) throws Exception{
		Thread.sleep(3000);
		((JavascriptExecutor) wd).executeScript("window.open()");
		Thread.sleep(2000);
		ArrayList<String> tabs1 = new ArrayList<String>(wd.getWindowHandles());
		Thread.sleep(2000);
		wd.close();
		wd.switchTo().window(tabs1.get(1));
		Thread.sleep(2000);
		wd.get(currentURL);
		Thread.sleep(5000);
		CloseCookies();
	}
		

	////////////NEW CODE HERE

	// Get Location
	public void getLocation(Row rowN, XSSFSheet sheet, WebDriverWait wait10, WebDriverWait wait20, WebDriverWait wait30, int i, Row rowHeading4, int HeadN, int s) throws IOException, InterruptedException {
	    try {
	        WebElement location = locateElement(
	            "//div[@class='accordionSection_carouselContainer__Y5UzM']//div/div[1]//div[2]/a",
	            "//a[@class='InThisTopic_location__VJhFE']"
	        );

	        if (location != null) {
	            // Element found, proceed with operations
	            System.out.println("Location " + i + ": " + location.getAttribute("innerText"));

	            // Save location details in Excel
	            rowN.createCell(5).setCellValue(location.getAttribute("innerText"));
	            rowN.createCell(6).setCellValue("Pass");
	            String linkUrl = location.getAttribute("href");
	            rowN.createCell(7).setCellValue(linkUrl);
	            System.out.println(linkUrl);

	            // Verify the location URL with an HTTP request
	            boolean isURLValid = verifyURL(linkUrl);

	            if (isURLValid) {
	                System.out.println("URL is valid.");
	                rowN.createCell(8).setCellValue("OK");

	                // Retrieve HREFs from the location URL
	                List<WebElement> hrefElements = wd.findElements(By.xpath("//link[@hreflang]"));

	                // Dynamically create HREF headers if not already created
	                createHREFHeaders(sheet, rowHeading4, hrefElements);

	                // Add HREF links to the current row
	                addHREFLinksToRow(rowN, hrefElements, sheet);
	            } else {
	                System.out.println("URL is not valid.");
	                rowN.createCell(8).setCellValue("Not Found");

	                // Highlight cell in red for invalid URL
	                highlightCellRed(sheet, rowN.getCell(8));
	            }
	        } else {
	            // Element not found with either xpath
	            System.out.println("Location not found.");
	            rowN.createCell(6).setCellValue("Fail");

	            // Highlight cell in red for location failure
	            highlightCellRed(sheet, rowN.getCell(6));
	        }
	    } catch (Exception e) {
	        System.out.println("Location " + i + ": Fail");
	        rowN.createCell(6).setCellValue("Fail");

	        // Highlight cell in red for location failure
	        highlightCellRed(sheet, rowN.getCell(6));
	    }
	}

	// Helper method to locate an element with fallback strategies
	private WebElement locateElement(String primaryXPath, String fallbackXPath) {
	    try {
	        return wd.findElement(By.xpath(primaryXPath));
	    } catch (NoSuchElementException e) {
	        try {
	            return wd.findElement(By.xpath(fallbackXPath));
	        } catch (NoSuchElementException e1) {
	            System.out.println("Element not found using either XPath.");
	            return null;
	        }
	    }
	}

	// Helper method to verify the URL with an HTTP request
	private boolean verifyURL(String linkUrl) {
	    try {
	        HttpURLConnection connection = (HttpURLConnection) new URL(linkUrl).openConnection();
	        connection.setRequestMethod("HEAD");
	        connection.setConnectTimeout(5000);
	        connection.setReadTimeout(5000);
	        int responseCode = connection.getResponseCode();
	        return responseCode >= 200 && responseCode < 300;
	    } catch (Exception e) {
	        System.out.println("Error verifying URL: " + e.getMessage());
	        return false;
	    }
	}

	private void createHREFHeaders(XSSFSheet sheet, Row rowHeading4, List<WebElement> hrefElements) {
	    if (hreflangColumnMap == null) {
	        hreflangColumnMap = new HashMap<>();
	    }

	    int HeadN = rowHeading4.getLastCellNum() > 0 ? rowHeading4.getLastCellNum() : 0; // Start at the next available column

	    for (WebElement hrefElement : hrefElements) {
	        String hreflang = hrefElement.getAttribute("hreflang");
	        if (hreflang != null) {
	            hreflang = normalizeHreflang(hreflang); // Normalize hreflang (e.g., "JA-JP" to "JA")

	            if (!hreflangColumnMap.containsKey(hreflang)) {
	                String header = hreflang.toUpperCase() + " HREF";
	                rowHeading4.createCell(HeadN).setCellValue(header);
	                hreflangColumnMap.put(hreflang, HeadN++);
	            }
	        }
	    }
	}

	// Helper to normalize hreflang (e.g., JA-JP -> JA)
	private String normalizeHreflang(String hreflang) {
	    if (hreflang.contains("-")) {
	        return hreflang.split("-")[0].toLowerCase();
	    }
	    return hreflang.toLowerCase();
	}


	private void addHREFLinksToRow(Row rowN, List<WebElement> hrefElements, XSSFSheet sheet) {
	    // Temporary map to track which HREFs were added
	    Map<String, String> hrefMap = new HashMap<>();
	    for (WebElement hrefElement : hrefElements) {
	        String hreflang = hrefElement.getAttribute("hreflang");
	        String href = hrefElement.getAttribute("href");

	        if (hreflang != null) {
	            hreflang = normalizeHreflang(hreflang); // Normalize hreflang
	            hrefMap.put(hreflang, href);
	        }
	    }

	    // Populate the Excel row
	    for (Map.Entry<String, Integer> entry : hreflangColumnMap.entrySet()) {
	        String hreflang = entry.getKey();
	        int columnIndex = entry.getValue();
	        Cell cell = rowN.createCell(columnIndex);

	        if (hrefMap.containsKey(hreflang)) {
	            cell.setCellValue(hrefMap.get(hreflang));
	        } else {
	            cell.setCellValue("NO HREF");

	            // Highlight the cell in yellow for missing HREFs
	            highlightCellYellow(sheet, cell);
	        }
	    }
	}





	// Add a helper to close popups
	private void closePopup() throws InterruptedException {
	    try {
	        wd.switchTo().defaultContent();
	        WebElement closeButton = wd.findElement(By.xpath(".//p[contains(@class,'modal_btnClose__eatUo') and contains(@class,'modal_headerElement__9HT5x')]"));
	        if (closeButton.isDisplayed()) {
	            closeButton.click();
	            System.out.println("Popup closed.");
	        } else {
	            System.out.println("Popup close button not visible.");
	        }
	    } catch (NoSuchElementException e) {
	        System.out.println("Popup close button not found.");
	    } catch (Exception e) {
	        System.out.println("Cannot Close Popup: " + e.getMessage());
	        Thread.sleep(2000);
	        ArrayList<String> tabs = new ArrayList<>(wd.getWindowHandles());
	        wd.switchTo().window(tabs.get(0));
	        Thread.sleep(2000);
	        wd.switchTo().defaultContent();
	        try {
	            wd.findElement(By.xpath(".//p[contains(@class,'modal_btnClose__eatUo') and contains(@class,'modal_headerElement__9HT5x')]")).click();
	        } catch (Exception innerEx) {
	            System.out.println("Failed to close popup after retry: " + innerEx.getMessage());
	        }
	    }
	}


	
	
	
	
	
}
