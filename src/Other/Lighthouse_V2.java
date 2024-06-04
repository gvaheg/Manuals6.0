package Other;

import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Lighthouse_V2 {

	public static void main(String[] args) throws Exception {
		// Setup Drivers and Browser
		WebDriver wd;
		XSSFWorkbook wb = new XSSFWorkbook();
		System.setProperty("webdriver.chrome.driver", "C:\\SeleniumDrivers\\chromedriver.exe");
		wd = new ChromeDriver();
		System.out.println("Browser Started!");
		wd.manage().window().maximize();
		wd.manage().timeouts().implicitlyWait(Duration.ofSeconds(500));
		WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(700));

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

		// ===== ALL URLs in Arrays =====
		String[] mainURLs = new String[69];
		String[] Versions = new String[69];
		// Home Pages
		
		mainURLs[0] = "https://web.dev/measure/?url=https%3A%2F%2Fwww.merckmanuals.com%2Fprofessional";
		Versions[0] = "EN PV Home Merck";
		
		mainURLs[1] = "https://web.dev/measure/?url=https%3A%2F%2Fwww.msdmanuals.com%2Fprofessional";
		Versions[1] = "EN PV Home";
		
		mainURLs[2] = "https://web.dev/measure/?url=https%3A%2F%2Fwww.msdmanuals.com%2Fde%2Fprofi";
		Versions[2] = "DE PV Home";
		
		mainURLs[3] = "https://web.dev/measure/?url=https%3A%2F%2Fwww.msdmanuals.com%2Fes%2Fprofessional";
		Versions[3] = "ES PV Home";
		
		mainURLs[4] = "https://web.dev/measure/?url=https%3A%2F%2Fwww.msdmanuals.com%2Ffr%2Fprofessional";
		Versions[4] = "FR PV Home";
		mainURLs[5] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fit%2Fprofessionale&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[5] = "IT PV Home";
		mainURLs[6] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fja-jp%2F%25E3%2583%2597%25E3%2583%25AD%25E3%2583%2595%25E3%2582%25A7%25E3%2583%2583%25E3%2582%25B7%25E3%2583%25A7%25E3%2583%258A%25E3%2583%25AB&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[6] = "JA PV Home";
		
		mainURLs[7] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fpt%2Fprofissional&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[7] = "PT PV Home";
		
		mainURLs[8] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fru%2F%25d0%25bf%25d1%2580%25d0%25be%25d1%2584%25d0%25b5%25d1%2581%25d1%2581%25d0%25b8%25d0%25be%25d0%25bd%25d0%25b0%25d0%25bb%25d1%258c%25d0%25bd%25d1%258b%25d0%25b9&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[8] = "RU PV Home";
			
		mainURLs[9] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.cn%2Fprofessional&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[9] = "CN PV Home";
		
		mainURLs[10] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fvi%2Fchuy%25c3%25aan-gia&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[10] = "VI PV Home";
		mainURLs[11] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.merckmanuals.com%2Fhome&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[11] = "EN CV Home Merck";
		mainURLs[12] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fhome&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[12] = "EN CV Home";
		
		mainURLs[13] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fde%2Fheim&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[13] = "DE CV Home";
		
		mainURLs[14] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fes%2Fhogar&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[14] = "ES CV Home";
		mainURLs[15] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Ffr%2Faccueil&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[15] = "FR CV Home";
		
		mainURLs[16] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fit%2Fcasa&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[16] = "IT CV Home";
		
		mainURLs[17] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fja-jp%2F%25e3%2583%259b%25e3%2583%25bc%25e3%2583%25a0&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[17] = "JA CV Home";
		
		mainURLs[18] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fko%2F%25ed%2599%2588&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[18] = "KO CV Home";
		
		mainURLs[19] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fpt%2Fcasa&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[19] = "PT CV Home";
		mainURLs[20] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fru%2F%25d0%25b4%25d0%25be%25d0%25bc%25d0%25b0&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[20] = "RU CV Home";
		
		mainURLs[21] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.cn%2Fhome&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[21] = "CN CV Home";
		
		mainURLs[22] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Far%2Fhome&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[22] = "AR CV Home";

		// Topic Pages ACNE
		mainURLs[23] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.merckmanuals.com%2Fprofessional%2Fdermatologic-disorders%2Facne-and-related-disorders%2Facne-vulgaris%3Fquery%3Dacne&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[23] = "EN PV Acne Merck";
		
		mainURLs[24] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fen-jp%2Fprofessional%2Fdermatologic-disorders%2Facne-and-related-disorders%2Facne-vulgaris&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[24] = "EN PV Acne";
		
		mainURLs[25] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fde%2Fprofi%2Ferkrankungen-der-haut%2Fakne-und-verwandte-erkrankungen%2Facne-vulgaris&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[25] = "DE PV Acne";
		mainURLs[26] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fes%2Fprofessional%2Ftrastornos-dermatol%25c3%25b3gicos%2Facn%25c3%25a9-y-trastornos-relacionados%2Facn%25c3%25a9-vulgar&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[26] = "ES PV Acne";
		
		mainURLs[27] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Ffr%2Fprofessional%2Ftroubles-dermatologiques%2Facn%25c3%25a9-et-pathologies-apparent%25c3%25a9es%2Facn%25c3%25a9-vulgaire&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[27] = "FR PV Acne";
		
		mainURLs[28] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fit%2Fprofessionale%2Fdisturbi-dermatologici%2Facne-e-disturbi-correlati%2Facne-vulgaris&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[28] = "IT PV Acne";
		mainURLs[29] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fja-jp%2F%25e3%2583%2597%25e3%2583%25ad%25e3%2583%2595%25e3%2582%25a7%25e3%2583%2583%25e3%2582%25b7%25e3%2583%25a7%25e3%2583%258a%25e3%2583%25ab%2F14-%25e7%259a%25ae%25e8%2586%259a%25e7%2596%25be%25e6%2582%25a3%2F%25e3%2581%2596%25e7%2598%25a1%25e3%2581%258a%25e3%2582%2588%25e3%2581%25b3%25e9%2596%25a2%25e9%2580%25a3%25e7%2596%25be%25e6%2582%25a3%2F%25e5%25b0%258b%25e5%25b8%25b8%25e6%2580%25a7%25e3%2581%2596%25e7%2598%25a1&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[29] = "JA PV Acne";
		mainURLs[30] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fpt%2Fprofissional%2Fdist%25c3%25barbios-dermatol%25c3%25b3gicos%2Facne-e-doen%25c3%25a7as-relacionadas%2Facne-vulgar&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[30] = "PT PV Acne";
		mainURLs[31] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fru%2F%25d0%25bf%25d1%2580%25d0%25be%25d1%2584%25d0%25b5%25d1%2581%25d1%2581%25d0%25b8%25d0%25be%25d0%25bd%25d0%25b0%25d0%25bb%25d1%258c%25d0%25bd%25d1%258b%25d0%25b9%2F%25d0%25b4%25d0%25b5%25d1%2580%25d0%25bc%25d0%25b0%25d1%2582%25d0%25be%25d0%25bb%25d0%25be%25d0%25b3%25d0%25b8%25d1%2587%25d0%25b5%25d1%2581%25d0%25ba%25d0%25b0%25d1%258f-%25d0%25bf%25d0%25b0%25d1%2582%25d0%25be%25d0%25bb%25d0%25be%25d0%25b3%25d0%25b8%25d1%258f%2F%25d0%25b0%25d0%25ba%25d0%25bd%25d0%25b5-%25d0%25b8-%25d1%2581%25d0%25b2%25d1%258f%25d0%25b7%25d0%25b0%25d0%25bd%25d0%25bd%25d1%258b%25d0%25b5-%25d1%2581-%25d0%25bd%25d0%25b8%25d0%25bc%25d0%25b8-%25d1%2580%25d0%25b0%25d1%2581%25d1%2581%25d1%2582%25d1%2580%25d0%25be%25d0%25b9%25d1%2581%25d1%2582%25d0%25b2%25d0%25b0%2F%25d0%25b2%25d1%2583%25d0%25bb%25d1%258c%25d0%25b3%25d0%25b0%25d1%2580%25d0%25bd%25d1%258b%25d0%25b5-%25d1%2583%25d0%25b3%25d1%2580%25d0%25b8&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[31] = "RU PV Acne";
		
		mainURLs[32] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.cn%2Fprofessional%2Fdermatologic-disorders%2Facne-and-related-disorders%2Facne-vulgaris&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[32] = "CN PV Acne";
		
		mainURLs[33] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fvi%2Fchuy%25c3%25aan-gia%2Fr%25e1%25bb%2591i-lo%25e1%25ba%25a1n-da-li%25e1%25bb%2585u%2Ftr%25e1%25bb%25a9ng-c%25c3%25a1-v%25c3%25a0-c%25c3%25a1c-b%25e1%25bb%2587nh-l%25c3%25bd-li%25c3%25aan-quan%2Ftr%25e1%25bb%25a9ng-c%25c3%25a1&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[33] = "VI PV Acne";
		mainURLs[34] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.merckmanuals.com%2Fhome%2Fskin-disorders%2Facne-and-related-disorders%2Facne%3Fquery%3Dacne&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[34] = "EN CV Acne Merck";
		mainURLs[35] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fhome%2Fskin-disorders%2Facne-and-related-disorders%2Facne&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[35] = "EN CV Acne";
		
		mainURLs[36] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fde%2Fheim%2Fhauterkrankungen%2Fakne-und-verwandte-erkrankungen%2Fakne&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[36] = "DE CV Acne";
		
		mainURLs[37] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fes%2Fhogar%2Ftrastornos-de-la-piel%2Facn%25c3%25a9-y-trastornos-relacionados%2Facn%25c3%25a9&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[37] = "ES CV Acne";
		mainURLs[38] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Ffr%2Faccueil%2Ftroubles-cutan%25c3%25a9s%2Facn%25c3%25a9-et-troubles-associ%25c3%25a9s%2Facn%25c3%25a9&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[38] = "FR CV Acne";
		mainURLs[39] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fit%2Fcasa%2Fpatologie-della-cute%2Facne-e-disturbi-correlati%2Facne&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[39] = "IT CV Acne";
		mainURLs[40] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fja-jp%2F%25E3%2583%259B%25E3%2583%25BC%25E3%2583%25A0%2F17-%25E7%259A%25AE%25E8%2586%259A%25E3%2581%25AE%25E7%2597%2585%25E6%25B0%2597%2F%25E3%2581%25AB%25E3%2581%258D%25E3%2581%25B3%25E3%2581%25A8%25E9%2596%25A2%25E9%2580%25A3%25E7%2596%25BE%25E6%2582%25A3%2F%25E3%2581%25AB%25E3%2581%258D%25E3%2581%25B3-%25E3%2581%2596%25E7%2598%25A1&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[40] = "JA CV Acne";
		mainURLs[41] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fko%2F%25ed%2599%2588%2F%25ed%2594%25bc%25eb%25b6%2580-%25ec%25a7%2588%25ed%2599%2598%2F%25ec%2597%25ac%25eb%2593%259c%25eb%25a6%2584-%25eb%25b0%258f-%25ea%25b4%2580%25eb%25a0%25a8-%25ec%25a7%2588%25ed%2599%2598%2F%25ec%2597%25ac%25eb%2593%259c%25eb%25a6%2584&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[41] = "KO CV Acne";
		mainURLs[42] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fpt%2Fcasa%2Fdist%25c3%25barbios-da-pele%2Facne-e-dist%25c3%25barbios-relacionados%2Facne&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[42] = "PT CV Acne";
		mainURLs[43] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fru%2F%25d0%25b4%25d0%25be%25d0%25bc%25d0%25b0%2F%25d0%25ba%25d0%25be%25d0%25b6%25d0%25bd%25d1%258b%25d0%25b5-%25d0%25b7%25d0%25b0%25d0%25b1%25d0%25be%25d0%25bb%25d0%25b5%25d0%25b2%25d0%25b0%25d0%25bd%25d0%25b8%25d1%258f%2F%25d1%2583%25d0%25b3%25d1%2580%25d0%25b5%25d0%25b2%25d0%25b0%25d1%258f-%25d1%2581%25d1%258b%25d0%25bf%25d1%258c-%25d0%25b8-%25d1%2580%25d0%25be%25d0%25b4%25d1%2581%25d1%2582%25d0%25b2%25d0%25b5%25d0%25bd%25d0%25bd%25d1%258b%25d0%25b5-%25d0%25b7%25d0%25b0%25d0%25b1%25d0%25be%25d0%25bb%25d0%25b5%25d0%25b2%25d0%25b0%25d0%25bd%25d0%25b8%25d1%258f%2F%25d1%2583%25d0%25b3%25d1%2580%25d0%25b5%25d0%25b2%25d0%25b0%25d1%258f-%25d1%2581%25d1%258b%25d0%25bf%25d1%258c&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[43] = "RU CV Acne";
		
		mainURLs[44] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.cn%2Fhome%2Fskin-disorders%2Facne-and-related-disorders%2Facne&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[44] = "CN CV Acne";
		mainURLs[45] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Far%2Fhome%2F%25d8%25a7%25d9%2584%25d8%25a7%25d8%25b6%25d8%25b7%25d8%25b1%25d8%25a7%25d8%25a8%25d8%25a7%25d8%25aa-%25d8%25a7%25d9%2584%25d8%25ac%25d9%2584%25d8%25af%25d9%258a%25d9%2591%25d9%258e%25d8%25a9%2F%25d8%25ad%25d8%25a8%25d9%2591-%25d8%25a7%25d9%2584%25d8%25b4%25d9%2591%25d9%258e%25d8%25a8%25d8%25a7%25d8%25a8-%25d9%2588%25d8%25a7%25d9%2584%25d8%25a7%25d8%25b6%25d8%25b7%25d8%25b1%25d8%25a7%25d8%25a8%25d8%25a7%25d8%25aa-%25d8%25b0%25d8%25a7%25d8%25aa-%25d8%25a7%25d9%2584%25d8%25b5%25d9%2591%25d9%2590%25d9%2584%25d8%25a9%2F%25d8%25ad%25d8%25a8-%25d8%25a7%25d9%2584%25d8%25b4%25d8%25a8%25d8%25a7%25d8%25a8-%25d8%25a7%25d9%2584%25d8%25b9%25d8%25af%25d9%2591&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[45] = "AR CV Acne";

		// Topic Pages Ankle Fractures
		mainURLs[46] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.merckmanuals.com%2Fprofessional%2Finjuries-poisoning%2Ffractures%2Fankle-fractures%3Fquery%3Dankle%2520fractures&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[46] = "EN PV Ankle Fractures Merck";
		mainURLs[47] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fprofessional%2Finjuries-poisoning%2Ffractures%2Fankle-fractures&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[47] = "EN PV Ankle Fractures";
		
		mainURLs[48] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fde%2Fprofi%2Fverletzungen%2C-vergiftungen%2Ffrakturen%2Fsprunggelenkfrakturen&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[48] = "DE PV Ankle Fractures";
		mainURLs[49] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fes%2Fprofessional%2Flesiones-y-envenenamientos%2Ffracturas%2Ffracturas-del-tobillo&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[49] = "ES PV Ankle Fractures";
		mainURLs[50] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Ffr%2Fprofessional%2Fblessures-empoisonnement%2Ffractures%2Ffractures-de-cheville&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[50] = "FR PV Ankle Fractures";
		mainURLs[51] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fit%2Fprofessionale%2Ftraumi-avvelenamento%2Ffratture%2Ffratture-alla-caviglia&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[51] = "IT PV Ankle Fractures";
		mainURLs[52] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fja-jp%2F%25e3%2583%2597%25e3%2583%25ad%25e3%2583%2595%25e3%2582%25a7%25e3%2583%2583%25e3%2582%25b7%25e3%2583%25a7%25e3%2583%258a%25e3%2583%25ab%2F22-%25e5%25a4%2596%25e5%2582%25b7%25e3%2581%25a8%25e4%25b8%25ad%25e6%25af%2592%2F%25e9%25aa%25a8%25e6%258a%2598-%25e8%2584%25b1%25e8%2587%25bc-%25e3%2581%258a%25e3%2582%2588%25e3%2581%25b3%25e6%258d%25bb%25e6%258c%25ab%2F%25e8%25b6%25b3%25e9%2596%25a2%25e7%25af%2580%25e9%25aa%25a8%25e6%258a%2598&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[52] = "JA PV Ankle Fractures";
		mainURLs[53] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fpt%2Fprofissional%2Fles%25c3%25b5es-intoxica%25c3%25a7%25c3%25a3o%2Ffraturas%2Ffraturas-do-tornozelo&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[53] = "PT PV Ankle Fractures";
		mainURLs[54] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fru%2F%25d0%25bf%25d1%2580%25d0%25be%25d1%2584%25d0%25b5%25d1%2581%25d1%2581%25d0%25b8%25d0%25be%25d0%25bd%25d0%25b0%25d0%25bb%25d1%258c%25d0%25bd%25d1%258b%25d0%25b9%2F%25d1%2582%25d1%2580%25d0%25b0%25d0%25b2%25d0%25bc%25d1%258b-%25d0%25be%25d1%2582%25d1%2580%25d0%25b0%25d0%25b2%25d0%25bb%25d0%25b5%25d0%25bd%25d0%25b8%25d1%258f%2F%25d0%25bf%25d0%25b5%25d1%2580%25d0%25b5%25d0%25bb%25d0%25be%25d0%25bc%25d1%258b%2F%25d0%25bf%25d0%25b5%25d1%2580%25d0%25b5%25d0%25bb%25d0%25be%25d0%25bc%25d1%258b-%25d0%25b3%25d0%25be%25d0%25bb%25d0%25b5%25d0%25bd%25d0%25be%25d1%2581%25d1%2582%25d0%25be%25d0%25bf%25d0%25bd%25d0%25be%25d0%25b3%25d0%25be-%25d1%2581%25d1%2583%25d1%2581%25d1%2582%25d0%25b0%25d0%25b2%25d0%25b0&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[54] = "RU PV Ankle Fractures";
		
		mainURLs[55] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.cn%2Fprofessional%2Finjuries-poisoning%2Ffractures%2Fankle-fractures&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[55] = "CN PV Ankle Fractures";
		
		mainURLs[56] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fvi%2Fchuy%25c3%25aan-gia%2Fch%25e1%25ba%25a5n-th%25c6%25b0%25c6%25a1ng-ng%25e1%25bb%2599-%25c4%2591%25e1%25bb%2599c%2Fg%25c3%25a3y-x%25c6%25b0%25c6%25a1ng%2Fg%25c3%25a3y-x%25c6%25b0%25c6%25a1ng-c%25e1%25bb%2595-ch%25c3%25a2n&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[56] = "VI PV Ankle Fractures";
		mainURLs[57] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.merckmanuals.com%2Fhome%2Finjuries-and-poisoning%2Ffractures%2Fankle-fractures%3Fquery%3Dankle%2520fractures&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[57] = "EN CV Ankle Fractures Merck";
		mainURLs[58] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fhome%2Finjuries-and-poisoning%2Ffractures%2Fankle-fractures%3Fquery%3Dankle%2520fractures&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[58] = "EN CV Ankle Fractures";
		mainURLs[59] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fde%2Fheim%2Fverletzungen-und-vergiftung%2Ffrakturen%2Fkn%25c3%25b6chelfrakturen&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[59] = "DE CV Ankle Fractures";
		mainURLs[60] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fes%2Fhogar%2Ftraumatismos-y-envenenamientos%2Ffracturas%2Ffracturas-de-tobillo&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[60] = "ES CV Ankle Fractures";
		mainURLs[61] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Ffr%2Faccueil%2Fl%25c3%25a9sions-et-intoxications%2Ffractures%2Ffractures-de-la-cheville&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[61] = "FR CV Ankle Fractures";
		mainURLs[62] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fit%2Fcasa%2Flesioni-e-avvelenamento%2Ffratture%2Ffratture-della-caviglia&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		
		Versions[62] = "IT CV Ankle Fractures";
		mainURLs[63] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fja-jp%2F%25e3%2583%259b%25e3%2583%25bc%25e3%2583%25a0%2F25-%25e5%25a4%2596%25e5%2582%25b7%25e3%2581%25a8%25e4%25b8%25ad%25e6%25af%2592%2F%25e9%25aa%25a8%25e6%258a%2598%2F%25e8%25b6%25b3%25e9%25a6%2596%25e3%2581%25ae%25e9%25aa%25a8%25e6%258a%2598&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[63] = "JA CV Ankle Fractures";
		mainURLs[64] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fko%2F%25ed%2599%2588%2F%25eb%25b6%2580%25ec%2583%2581-%25eb%25b0%258f-%25ec%25a4%2591%25eb%258f%2585%2F%25ea%25b3%25a8%25ec%25a0%2588%2F%25eb%25b0%259c%25eb%25aa%25a9-%25ea%25b3%25a8%25ec%25a0%2588&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[64] = "KO CV Ankle Fractures";
		
		mainURLs[65] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fpt%2Fcasa%2Fles%25c3%25b5es-e-envenenamentos%2Ffraturas%2Ffraturas-do-tornozelo&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[65] = "PT CV Ankle Fractures";
		
		mainURLs[66] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Fru%2F%25d0%25b4%25d0%25be%25d0%25bc%25d0%25b0%2F%25d1%2582%25d1%2580%25d0%25b0%25d0%25b2%25d0%25bc%25d1%258b-%25d0%25b8-%25d0%25be%25d1%2582%25d1%2580%25d0%25b0%25d0%25b2%25d0%25bb%25d0%25b5%25d0%25bd%25d0%25b8%25d1%258f%2F%25d0%25bf%25d0%25b5%25d1%2580%25d0%25b5%25d0%25bb%25d0%25be%25d0%25bc%25d1%258b%2F%25d0%25bf%25d0%25b5%25d1%2580%25d0%25b5%25d0%25bb%25d0%25be%25d0%25bc%25d1%258b-%25d0%25bb%25d0%25be%25d0%25b4%25d1%258b%25d0%25b6%25d0%25ba%25d0%25b8&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[66] = "RU CV Ankle Fractures";
		
		mainURLs[67] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.cn%2Fhome%2Finjuries-and-poisoning%2Ffractures%2Fankle-fractures&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[67] = "CN CV Ankle Fractures";
		mainURLs[68] = "https://googlechrome.github.io/lighthouse/viewer/?psiurl=https%3A%2F%2Fwww.msdmanuals.com%2Far%2Fhome%2F%25d8%25a7%25d9%2584%25d8%25a5%25d8%25b5%25d8%25a7%25d8%25a8%25d8%25a7%25d8%25aa-%25d9%2588%25d8%25a7%25d9%2584%25d8%25aa%25d9%2591%25d9%258e%25d8%25b3%25d9%2585%25d9%2591%25d9%258f%25d9%2585%2F%25d8%25ad%25d8%25a7%25d9%2584%25d8%25a7%25d8%25aa%25d9%258f-%25d8%25a7%25d9%2584%25d9%2583%25d8%25b3%25d9%2588%25d8%25b1-%25d9%2588%25d8%25a7%25d9%2584%25d8%25ae%25d9%2584%25d9%2588%25d8%25b9-%25d9%2588%25d8%25a7%25d9%2584%25d8%25a7%25d9%2584%25d8%25aa%25d9%2588%25d8%25a7%25d8%25a1%25d8%25a7%25d8%25aa%2F%25d9%2583%25d9%258f%25d8%25b3%25d9%2588%25d8%25b1%25d9%258f-%25d8%25a7%25d9%2584%25d9%2583%25d8%25a7%25d8%25ad%25d9%2584&strategy=mobile&category=performance&category=accessibility&category=best-practices&category=seo&category=pwa&utm_source=lh-chrome-ext";
		Versions[68] = "AR CV Ankle Fractures";

		// Main Loop for HOME PAGES
		for (int j = 0; j < 69; j++) {
			try {
				Row rowN = sheet.createRow(rowNum++);
				// Navigate to version
				wd.get(mainURLs[j]);
				//Thread.sleep(7000);
				String currentURL = wd.getCurrentUrl();
				Thread.sleep(4000);
				wd.findElement(By.xpath("//*[@id=\"run-lh-button\"]")).click();
				try {
					wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[2]/a[1]/div[@class=\"lh-gauge__percentage\"]")));
	
				} catch (Exception e) {
					((JavascriptExecutor) wd).executeScript("window.open()");
					Thread.sleep(1000);
					ArrayList<String> tabs = new ArrayList<String>(wd.getWindowHandles());
					wd.close();
					Thread.sleep(2000);
					wd.switchTo().window(tabs.get(1));
					wd.get(currentURL);
				}
				
				System.out.println(Versions[j]);
				rowN.createCell(0).setCellValue(Versions[j]);
				// performance
				WebElement performance = wd
						.findElement(By.xpath(".//div[2]/a[1]/div[2]"));
				System.out.println("performance: " + performance.getAttribute("innerHTML"));
				rowN.createCell(1).setCellValue(performance.getAttribute("innerHTML"));
				//Thread.sleep(500);
				// accessibility
				WebElement accessibility = wd
						.findElement(By.xpath(".//div[2]/a[2]/div[2]"));
				System.out.println("accessibility: " + accessibility.getAttribute("innerHTML"));
				rowN.createCell(2).setCellValue(accessibility.getAttribute("innerHTML"));
				//Thread.sleep(500);
				// best-practices
				WebElement bestpractices = wd
						.findElement(By.xpath(".//div[2]/a[3]/div[2]"));
				System.out.println("best-practices: " + bestpractices.getAttribute("innerHTML"));
				rowN.createCell(3).setCellValue(bestpractices.getAttribute("innerHTML"));
				//Thread.sleep(500);
				// seo
				WebElement seo = wd.findElement(By.xpath(".//div[2]/a[4]/div[2]"));
				System.out.println("seo: " + seo.getAttribute("innerHTML"));
				rowN.createCell(4).setCellValue(seo.getAttribute("innerHTML"));
				//Thread.sleep(500);

				// first-contentful-paint
				WebElement firstcontentfulpaint = wd
						.findElement(By.xpath("//*[@id=\"first-contentful-paint\"]/div/div[2]"));
				System.out.println("first-contentful-paint: " + firstcontentfulpaint.getText());
				rowN.createCell(5).setCellValue(firstcontentfulpaint.getText());
				//Thread.sleep(500);
				// speed-index
				WebElement speedindex = wd
						.findElement(By.xpath("//*[@id=\"speed-index\"]/div/div[2]"));
				System.out.println("speed-index: " + speedindex.getText());
				rowN.createCell(6).setCellValue(speedindex.getText());
				//Thread.sleep(500);
				// largest-contentful-paint
				WebElement largestcontentfulpaint = wd
						.findElement(By.xpath("//*[@id=\"largest-contentful-paint\"]/div/div[2]"));
				System.out.println("largest-contentful-paint: " + largestcontentfulpaint.getText());
				rowN.createCell(7).setCellValue(largestcontentfulpaint.getText());
				//Thread.sleep(500);
				// interactive
				WebElement interactive = wd
						.findElement(By.xpath("//*[@id=\"interactive\"]/div/div[2]"));
				System.out.println("interactive: " + interactive.getText());
				rowN.createCell(8).setCellValue(interactive.getText());
				//Thread.sleep(500);
				// total-blocking-time
				WebElement totalblockingtime = wd
						.findElement(By.xpath("//*[@id=\"total-blocking-time\"]/div/div[2]"));
				System.out.println("total-blocking-time: " + totalblockingtime.getText());
				rowN.createCell(9).setCellValue(totalblockingtime.getText());
				//Thread.sleep(500);
				// cumulative-layout-shift
				WebElement cumulativelayoutshift = wd
						.findElement(By.xpath("//*[@id=\"cumulative-layout-shift\"]/div/div[2]"));
				System.out.println("cumulative-layout-shift: " + cumulativelayoutshift.getText());
				rowN.createCell(10).setCellValue(cumulativelayoutshift.getText());
				//Thread.sleep(500);

				FileOutputStream fout = new FileOutputStream("C:\\users\\gvahe\\Desktop\\Lighthouse2.xlsx");
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
				// wd.get(currentURL);
				Thread.sleep(10000);
			}

		}

		// Close excel and browser
		wb.close();
		wd.close();
		System.out.println("Browser Closed!");
	}

}
