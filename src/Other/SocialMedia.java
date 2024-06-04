package Other;

import java.io.FileOutputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

public class SocialMedia {

	public static void main(String[] args) throws Exception {
		// Setup Drivers and Browser
		WebDriver wd;
		XSSFWorkbook wb = new XSSFWorkbook();
		System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
		wd = new FirefoxDriver();
		System.out.println("Browser Started!");
		wd.manage().window().maximize();
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(3));
		WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(3));

		// ===== ALL URLs in Arrays =====
		String[] mainURLs = new String[85];
		String[] Versions = new String[85];

		
		// English (EN-US)
		mainURLs[0] = "https://east.merckmanuals.com";
		Versions[0] = "English (EN-US) PV - Landing Page";
		mainURLs[1] = "https://east.merckmanuals.com/professional";
		Versions[1] = "English (EN-US) PV - Home Page";
		mainURLs[2] = "https://east.merckmanuals.com/home";
		Versions[2] = "English (EN-US) CV - Home Page";
		mainURLs[3] = "https://east.merckmanuals.com/home";
		Versions[3] = "English (EN-US) CV - Topic Page";
		// Spanish (ES)
		mainURLs[4] = "https://east.msdmanuals.com/es";
		Versions[4] = "Spanish (ES) - Landing Page";
		mainURLs[5] = "https://east.msdmanuals.com/es/professional";
		Versions[5] = "Spanish (ES) PV - Home Page";
		mainURLs[6] = "https://east.msdmanuals.com/es/hogar";
		Versions[6] = "Spanish (ES) CV - Home Page";
		// German (DE)
		mainURLs[7] = "https://east.msdmanuals.com/de";
		Versions[7] = "German (DE) - Landing Page";
		mainURLs[8] = "https://east.msdmanuals.com/de/profi";
		Versions[8] = "German (DE) PV - Home Page";
		mainURLs[9] = "https://east.msdmanuals.com/de/heim";
		Versions[9] = "German (DE) CV - Home Page";
		// French (FR)
		mainURLs[10] = "https://east.msdmanuals.com/fr";
		Versions[10] = "French (FR) - Landing Page";
		mainURLs[11] = "https://east.msdmanuals.com/fr/professional";
		Versions[11] = "French (FR) PV - Home Page";
		mainURLs[12] = "https://east.msdmanuals.com/fr/accueil";
		Versions[12] = "French (FR) CV - Home Page";
		// Italian (IT)
		mainURLs[13] = "https://east.msdmanuals.com/it";
		Versions[13] = "Italian (IT) - Landing Page";
		mainURLs[14] = "https://east.msdmanuals.com/it/professionale";
		Versions[14] = "Italian (IT) PV - Home Page";
		mainURLs[15] = "https://east.msdmanuals.com/it/casa";
		Versions[15] = "Italian (IT) CV - Home Page";
		// Japanese (JA)
		mainURLs[16] = "https://east.msdmanuals.com/ja";
		Versions[16] = "Japanese (JA) - Landing Page";
		mainURLs[17] = "https://east.msdmanuals.com/ja/%E3%83%97%E3%83%AD%E3%83%95%E3%82%A7%E3%83%83%E3%82%B7%E3%83%A7%E3%83%8A%E3%83%AB";
		Versions[17] = "Japanese (JA) PV - Home Page";
		mainURLs[18] = "https://east.msdmanuals.com/ja/%E3%83%9B%E3%83%BC%E3%83%A0";
		Versions[18] = "Japanese (JA) CV - Home Page";
		// Korean (KO)
		mainURLs[19] = "https://east.msdmanuals.com/ko";
		Versions[19] = "Korean (KO) - Landing Page";
		mainURLs[20] = "https://east.msdmanuals.com/en-kr/professional";
		Versions[20] = "Korean (KO) PV - Home Page";
		mainURLs[21] = "https://east.msdmanuals.com/ko/%ED%99%88";
		Versions[21] = "Korean (KO) CV - Home Page";
		// Portuguese (PT)
		mainURLs[22] = "https://east.msdmanuals.com/pt";
		Versions[22] = "Portuguese (PT) - Landing Page";
		mainURLs[23] = "https://east.msdmanuals.com/pt/profissional";
		Versions[23] = "Portuguese (PT) PV - Home Page";
		mainURLs[24] = "https://east.msdmanuals.com/pt/casa";
		Versions[24] = "Portuguese (PT) CV - Home Page";
		// Russian (RU)
		mainURLs[25] = "https://east.msdmanuals.com/ru";
		Versions[25] = "Russian (RU) - Landing Page";
		mainURLs[26] = "https://east.msdmanuals.com/ru/%D0%BF%D1%80%D0%BE%D1%84%D0%B5%D1%81%D1%81%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9";
		Versions[26] = "Russian (RU) PV - Home Page";
		mainURLs[27] = "https://east.msdmanuals.com/ru/%D0%B4%D0%BE%D0%BC%D0%B0";
		Versions[27] = "Russian (RU) CV - Home Page";		
		// Chinese (ZH)
		mainURLs[28] = "https://east.msdmanuals.com/zh";
		Versions[28] = "Chinese (ZH) - Landing Page";
		mainURLs[29] = "https://east.msdmanuals.com/zh/%E4%B8%93%E4%B8%9A";
		Versions[29] = "Chinese (ZH) PV - Home Page";
		mainURLs[30] = "https://east.msdmanuals.com/zh/%E9%A6%96%E9%A1%B5";
		Versions[30] = "Chinese (ZH) CV - Home Page";
		// Vet
		mainURLs[31] = "https://east.msdvetmanual.com";
		Versions[31] = "MSD Vet Manual - Home Page";
		mainURLs[32] = "https://east.merckvetmanual.com";
		Versions[32] = "MERCK Vet Manual - Home Page";
		// English EN
		mainURLs[33] = "https://east.msdmanuals.com/";
		Versions[33] = "English (EN) PV - Landing Page";
		mainURLs[34] = "https://east.msdmanuals.com/professional";
		Versions[34] = "English (EN) PV - Home Page";
		mainURLs[35] = "https://east.msdmanuals.com/home";
		Versions[35] = "English (EN) CV - Home Page";
		// English (EN-AU)
		mainURLs[36] = "https://east.msdmanuals.com/en-au/";
		Versions[36] = "English (EN-AU) PV - Landing Page";
		mainURLs[37] = "https://east.msdmanuals.com/en-au/professional";
		Versions[37] = "English (EN-AU) PV - Home Page";
		mainURLs[38] = "https://east.msdmanuals.com/en-au/home";
		Versions[38] = "English (EN-AU) CV - Home Page";
		// English (EN-CA)
		mainURLs[39] = "https://east.merckmanuals.com/en-ca/";
		Versions[39] = "English (EN-CA) PV - Landing Page";
		mainURLs[40] = "https://east.merckmanuals.com/en-ca/professional";
		Versions[40] = "English (EN-CA) PV - Home Page";
		mainURLs[41] = "https://east.merckmanuals.com/en-ca/home";
		Versions[41] = "English (EN-CA) CV - Home Page";
		// English (EN-GB)
		mainURLs[42] = "https://east.msdmanuals.com/en-gb/";
		Versions[42] = "English (EN-GB) PV - Landing Page";
		mainURLs[43] = "https://east.msdmanuals.com/en-gb/professional";
		Versions[43] = "English (EN-GB) PV - Home Page";
		mainURLs[44] = "https://east.msdmanuals.com/en-gb/home";
		Versions[44] = "English (EN-GB) CV - Home Page";
		// English (EN-KR)
		mainURLs[45] = "https://east.msdmanuals.com/en-kr/";
		Versions[45] = "English (EN-KR) PV - Landing Page";
		mainURLs[46] = "https://east.msdmanuals.com/en-kr/professional";
		Versions[46] = "English (EN-KR) PV - Home Page";
		mainURLs[47] = "https://east.msdmanuals.com/en-kr/home";
		Versions[47] = "English (EN-KR) CV - Home Page";
		// English (EN-NZ)
		mainURLs[48] = "https://east.msdmanuals.com/en-nz/";
		Versions[48] = "English (EN-NZ) PV - Landing Page";
		mainURLs[49] = "https://east.msdmanuals.com/en-nz/professional";
		Versions[49] = "English (EN-NZ) PV - Home Page";
		mainURLs[50] = "https://east.msdmanuals.com/en-nz/home";
		Versions[50] = "English (EN-NZ) CV - Home Page";
		// English (EN-PR)
		mainURLs[51] = "https://east.merckmanuals.com/en-pr/";
		Versions[51] = "English (EN-PR) PV - Landing Page";
		mainURLs[52] = "https://east.merckmanuals.com/en-pr/professional";
		Versions[52] = "English (EN-PR) PV - Home Page";
		mainURLs[53] = "https://east.merckmanuals.com/en-pr/home";
		Versions[53] = "English (EN-PR) CV - Home Page";
		// English (EN-PT)
		mainURLs[54] = "https://east.msdmanuals.com/en-pt/";
		Versions[54] = "English (EN-PT) PV - Landing Page";
		mainURLs[55] = "https://east.msdmanuals.com/en-pt/professional";
		Versions[55] = "English (EN-PT) PV - Home Page";
		mainURLs[56] = "https://east.msdmanuals.com/en-pt/home";
		Versions[56] = "English (EN-PT) CV - Home Page";
		// English (EN-SG)
		mainURLs[57] = "https://east.msdmanuals.com/en-sg/";
		Versions[57] = "English (EN-SG) PV - Landing Page";
		mainURLs[58] = "https://east.msdmanuals.com/en-sg/professional";
		Versions[58] = "English (EN-SG) PV - Home Page";
		mainURLs[59] = "https://east.msdmanuals.com/en-sg/home";
		Versions[59] = "English (EN-SG) CV - Home Page";
		// English (FR-CA)
		mainURLs[60] = "https://east.merckmanuals.com/fr-ca";
		Versions[60] = "English (FR-CA) PV - Landing Page";
		mainURLs[61] = "https://east.merckmanuals.com/fr-ca/professional";
		Versions[61] = "English (FR-CA) PV - Home Page";
		mainURLs[62] = "https://east.merckmanuals.com/fr-ca/accueil";
		Versions[62] = "English (FR-CA) CV - Home Page";
		// English (ES-US)
		mainURLs[63] = "https://east.merckmanuals.com/es-us";
		Versions[63] = "English (ES-US) PV - Landing Page";
		mainURLs[64] = "https://east.merckmanuals.com/es-us/professional";
		Versions[64] = "English (ES-US) PV - Home Page";
		mainURLs[65] = "https://east.merckmanuals.com/es-us/home";
		Versions[65] = "English (ES-US) CV - Home Page";
		// Spanish (ES-CL)
		mainURLs[66] = "https://east.msdmanuals.com/es-cl";
		Versions[66] = "Spanish (ES-CL) - Landing Page";
		mainURLs[67] = "https://east.msdmanuals.com/es-cl/professional";
		Versions[67] = "Spanish (ES-CL) PV - Home Page";
		mainURLs[68] = "https://east.msdmanuals.com/es-cl/hogar";
		Versions[68] = "Spanish (ES-CL) CV - Home Page";
		// Spanish (ES-CO)
		mainURLs[69] = "https://east.msdmanuals.com/es-co";
		Versions[69] = "Spanish (ES-CO) - Landing Page";
		mainURLs[70] = "https://east.msdmanuals.com/es-co/professional";
		Versions[70] = "Spanish (ES-CO) PV - Home Page";
		mainURLs[71] = "https://east.msdmanuals.com/es-co/hogar";
		Versions[71] = "Spanish (ES-CO) CV - Home Page";
		// Spanish (ES-CR)
		mainURLs[72] = "https://east.msdmanuals.com/es-cr";
		Versions[72] = "Spanish (ES-CR) - Landing Page";
		mainURLs[73] = "https://east.msdmanuals.com/es-cr/professional";
		Versions[73] = "Spanish (ES-CR) PV - Home Page";
		mainURLs[74] = "https://east.msdmanuals.com/es-cr/hogar";
		Versions[74] = "Spanish (ES-CR) CV - Home Page";
		// Spanish (ES-DO)
		mainURLs[75] = "https://east.msdmanuals.com/es-do";
		Versions[75] = "Spanish (ES-DO) - Landing Page";
		mainURLs[76] = "https://east.msdmanuals.com/es-do/professional";
		Versions[76] = "Spanish (ES-DO) PV - Home Page";
		mainURLs[77] = "https://east.msdmanuals.com/es-do/hogar";
		Versions[77] = "Spanish (ES-DO) CV - Home Page";
		// Spanish (ES-EC)
		mainURLs[78] = "https://east.msdmanuals.com/es-ec";
		Versions[78] = "Spanish (ES-EC) - Landing Page";
		mainURLs[79] = "https://east.msdmanuals.com/es-ec/professional";
		Versions[79] = "Spanish (ES-EC) PV - Home Page";
		mainURLs[80] = "https://east.msdmanuals.com/es-ec/hogar";
		Versions[80] = "Spanish (ES-EC) CV - Home Page";
		// Vietnamese (VI)
		mainURLs[81] = "https://east.msdmanuals.com/vi";
		Versions[81] = "Vietnamese (VI) - Landing Page";
		mainURLs[82] = "https://east.msdmanuals.com/vi/chuy%C3%AAn-gia";
		Versions[82] = "Vietnamese (VI) PV - Home Page";
		
		// Arabic (AR)
		mainURLs[83] = "https://www.msdmanuals.com/ar";
		Versions[83] = "Arabic (AR) CV - Landing Page";
		mainURLs[84] = "https://www.msdmanuals.com/ar/home";
		Versions[84] = "Arabic (AR) CV - Home Page";
		
		
		// Excel Create Headings
		XSSFSheet sheet = wb.createSheet();
		Row rowHeading0 = sheet.createRow(0);
		rowHeading0.createCell(0).setCellValue("VERSION");
		rowHeading0.createCell(1).setCellValue("VERSION URL");
		rowHeading0.createCell(2).setCellValue("Social Media");
		rowHeading0.createCell(3).setCellValue("Social Media URL");

		int rowNum = 1;
		// Main Loop
		for (int j = 0; j < 85; j++) {
			
			try {
				// Navigate to version
				wd.get(mainURLs[j]);
				// Close Cookies
				try {
					WebElement AcceptCookies = wd.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
					AcceptCookies.click();
				} catch (Exception e) {
					System.out.println("Can't Close Cookies");
				}
				// End Close Cookies
				Thread.sleep(5000);
				System.out.println(Versions[j]);
				
				
				// Social Media Links how many
				int SocialCount = wd.findElements(By.xpath(".//footer/div/div[3]/div[1]/div[1]/div/a")).size();
				System.out.println("Social Media on this page: " + SocialCount);
				int z = SocialCount;
							if (z>0) {
								for (int i = 1; i < SocialCount + 1; i++) {
									Row rowN = sheet.createRow(rowNum++);
									rowN.createCell(0).setCellValue(Versions[j]);
									rowN.createCell(1).setCellValue(mainURLs[j]);
									try {
										WebElement Social = wd.findElement(By.xpath(".//footer/div/div[3]/div[1]/div[1]/div/a["+i+"]"));
										System.out.println(Social.getText());
										rowN.createCell(2).setCellValue(Social.getText());
										String linkUrl = Social.getAttribute("href");
										System.out.println("Social URL: " + linkUrl);
										rowN.createCell(3).setCellValue(linkUrl);
									} catch (Exception e) {
										System.out.println("Can't Find Social");
									}
							}
							}else {
								Row rowN = sheet.createRow(rowNum++);
								rowN.createCell(0).setCellValue(Versions[j]);
								rowN.createCell(1).setCellValue(mainURLs[j]);
								rowN.createCell(2).setCellValue("No Social Media");
							}		
					
				

				FileOutputStream fout = new FileOutputStream("C:\\users\\gvahe\\Desktop\\SocialMedia.xlsx");
				wb.write(fout);
				fout.close();
				System.out.println("Written Successfully!");
				
			} catch (Exception e) {
				System.out.println("ERROR! Web page is not responding. Relopening the page... ");
				((JavascriptExecutor) wd).executeScript("window.open()");
				Thread.sleep(1000);
				ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
				wd.close();
				Thread.sleep(2000);
				wd.switchTo().window(tabs.get(1));
				//wd.get(currentURL);
				Thread.sleep(10000);
			}
		}
		// Close excel and browser
		wb.close();
		wd.close();
		System.out.println("Browser Closed!");
	}

}
