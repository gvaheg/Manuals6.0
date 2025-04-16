package ResourcesNEW;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.By;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import java.time.Duration;

public class Calculators extends All_Functions {
	@SuppressWarnings("deprecation")
	private void setupWebDriver() {
	    FirefoxOptions options = new FirefoxOptions();
	    options.setHeadless(true); // Enable headless mode
	    wd = new FirefoxDriver(options); // Initialize WebDriver with options
	}
    private WebDriver wd;
    @BeforeClass
    public void startBrowser() throws Exception {
        wb = new XSSFWorkbook();
        System.out.println("Browser Started!");
        getDate();
    }
    @DataProvider(name = "languageVersions")
    public Object[][] languageVersions() {

        return new Object[][] {
            // URL, Label, NeedsSpecialHandling
        	
            {"https://www.merckmanuals.com/professional/pages-with-widgets/clinical-calculators?mode=list&order=bysection&order=bysection", "PROD EN-US PV", false},
            {"https://www.msdmanuals.com/pt/profissional/pages-with-widgets/calculadoras-cl%C3%ADnicas?mode=list&order=bysection", "PROD PT PV", false},
            {"https://www.msdmanuals.com/ja-jp/professional/pages-with-widgets/%E5%8C%BB%E5%AD%A6%E8%A8%88%E7%AE%97%E3%83%84%E3%83%BC%E3%83%AB%EF%BC%88%E5%AD%A6%E7%BF%92%E7%94%A8%EF%BC%89?mode=list&order=bysection", "PROD JA PV", false},
            {"https://www.msdmanuals.com/fr/professional/pages-with-widgets/calculateurs-cliniques?mode=list&order=bysection", "PROD FR PV", false},
            {"https://www.msdmanuals.com/es/professional/pages-with-widgets/calculadoras-cl%C3%ADnicas?mode=list&order=bysection", "PROD ES PV", false},
            {"https://www.msdmanuals.com/de/profi/pages-with-widgets/klinische-rechner?mode=list&order=bysection", "PROD DE PV", false},
             
            {"https://www.msdmanuals.com/it/professionale/pages-with-widgets/calcolatori-clinici?mode=list&order=bysection", "PROD IT PV", false},
            {"https://www.msdmanuals.com/ru/professional/pages-with-widgets/%D0%BA%D0%BB%D0%B8%D0%BD%D0%B8%D1%87%D0%B5%D1%81%D0%BA%D0%B8%D0%B5-%D0%BA%D0%B0%D0%BB%D1%8C%D0%BA%D1%83%D0%BB%D1%8F%D1%82%D0%BE%D1%80%D1%8B?mode=list&order=bysection", "PROD RU PV", true},
          
            {"https://www.msdmanuals.cn/professional/pages-with-widgets/clinical-calculators?mode=list&order=bysection", "PROD CN PV", true},
           
            
            {"https://www.merckvetmanual.com/pages-with-widgets/clinical-calculators?mode=list&order=bysection", "PROD MM VET", false},
            {"https://www.msdvetmanual.com/pages-with-widgets/clinical-calculators?mode=list&order=bysection", "PROD MSD VET", false},
            {"https://www.msdmanuals.com/professional/pages-with-widgets/clinical-calculators?mode=list&order=bysection", "PROD EN MSD PV", false},
            {"https://www.msdmanuals.com/vi/professional/pages-with-widgets/clinical-calculators?mode=list&order=bysection", "PROD VI PV", false},
            {"https://www.msdmanuals.com/uk/professional/pages-with-widgets/clinical-calculators?mode=list&order=bysection", "PROD UK PV", false}

        };

    }


    @Test(dataProvider = "languageVersions")
    public void verifyLanguageVersion(String url, String version, boolean hasPopup) {
        try {
            System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
            setupWebDriver();
            All_Functions.wd = wd; // Set the WebDriver instance in All_Functions
            System.out.println("Navigating to URL: " + url);
            wd.get(url);
            WebDriverWait wait = new WebDriverWait(wd, Duration.ofSeconds(10));
            wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("body")));
            System.out.println("Page loaded successfully.");
            if (hasPopup) {
                handlePopups();
            }
            System.out.println("VERSION: " + version);
            verifyCalculators();
        } catch (Exception e) {
            System.out.println("Page Error for version: " + version);
            e.printStackTrace();
        } finally {
            if (wd != null) {
                wd.close();
                System.out.println("Browser closed.");
            }
        }
    }



    @AfterClass
    public void closeBrowser() throws Exception {
        getDate();
        System.out.println("================Browser Closed!================");
    }
}