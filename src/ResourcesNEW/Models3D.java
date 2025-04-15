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

public class Models3D extends All_Functions {
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
            {"https://www.merckmanuals.com/professional/pages-with-widgets/3d-models?mode=list", "PROD EN-US PV", false},
            {"https://www.merckmanuals.com/home/pages-with-widgets/3d-models?mode=list", "PROD EN-US CV", false},
            {"https://www.msdmanuals.com/pt/profissional/pages-with-widgets/modelos-3d?mode=list", "PROD PT PV", false},
            {"https://www.msdmanuals.com/pt/casa/pages-with-widgets/modelos-3d?mode=list", "PROD PT CV", false},
            {"https://www.msdmanuals.com/ja-jp/professional/pages-with-widgets/3d-models?mode=list", "PROD JA PV", false},
            {"https://www.msdmanuals.com/ja-jp/home/pages-with-widgets/3d-models?mode=list", "PROD JA CV", false},
            {"https://www.msdmanuals.com/fr/professional/pages-with-widgets/mod%C3%A8les-3d?mode=list", "PROD FR PV", false},
            {"https://www.msdmanuals.com/fr/accueil/pages-with-widgets/mod%C3%A8les-3d?mode=list", "PROD FR CV", false},
            {"https://www.msdmanuals.com/es/professional/pages-with-widgets/modelos-3d?mode=list", "PROD ES PV", false},
            {"https://www.msdmanuals.com/es/hogar/pages-with-widgets/modelos-3d?mode=list", "PROD ES CV", false},
            {"https://www.msdmanuals.com/de/profi/pages-with-widgets/3d-modelle?mode=list", "PROD DE PV", false},
            {"https://www.msdmanuals.com/de/heim/pages-with-widgets/3d-modelle?mode=list", "PROD DE CV", false},
            {"https://www.msdmanuals.com/it/professionale/pages-with-widgets/modelli-3d?mode=list", "PROD IT PV", false},
            {"https://www.msdmanuals.com/it/casa/pages-with-widgets/modelli-3d?mode=list", "PROD IT CV", false},
            {"https://www.msdmanuals.com/ru/professional/pages-with-widgets/3d-%EB%AA%A8%EB%8D%B8?mode=list", "PROD RU PV", false},
            {"https://www.msdmanuals.com/ru/home/pages-with-widgets/3d-%D0%BC%D0%BE%D0%B4%D0%B5%D0%BB%D0%B8?mode=list", "PROD RU CV", true},
            {"https://www.msdmanuals.cn/professional/pages-with-widgets/3d-models?mode=list", "PROD CN PV", true},
            {"https://www.msdmanuals.cn/home/pages-with-widgets/biodigital?mode=list", "PROD CN CV", true},
            {"https://www.msdmanuals.com/ko/home/pages-with-widgets/3d-%EB%AA%A8%EB%8D%B8?mode=list", "PROD KO CV", false},
            {"https://www.msdmanuals.com/professional/pages-with-widgets/3d-models?mode=list", "PROD EN MSD PV", false},
            {"https://www.msdmanuals.com/home/pages-with-widgets/3d-models?mode=list", "PROD EN MSD CV", false},
            {"https://www.msdmanuals.com/ar/home/pages-with-widgets/%D9%86%D9%85%D8%A7%D8%B0%D8%AC-%D8%AB%D9%84%D8%A7%D8%AB%D9%8A%D8%A9-%D8%A7%D9%84%D8%A3%D8%A8%D8%B9%D8%A7%D8%AF?mode=list", "PROD AR CV", false},
            {"https://www.msdmanuals.com/vi/professional/pages-with-widgets/c%C3%A1c-m%C3%B4-h%C3%ACnh-3d?mode=list", "PROD VI PV", false},
            {"https://www.msdmanuals.com/uk/professional/pages-with-widgets/%D1%82%D1%80%D0%B8%D0%B2%D0%B8%D0%BC%D1%96%D1%80%D0%BD%D1%96-%D0%BC%D0%BE%D0%B4%D0%B5%D0%BB%D1%96?mode=list", "PROD UK PV", false},
            {"https://www.msdmanuals.com/hi/home/pages-with-widgets/biodigital?mode=list", "PROD HI CV", false},
            {"https://www.msdmanuals.com/sw/home/pages-with-widgets/3d-models?mode=list", "PROD SW CV", false}

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
            verifyModels3D();
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