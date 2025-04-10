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

public class Figures extends All_Functions {
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
            {"https://www.merckmanuals.com/professional/pages-with-widgets/figures?mode=list", "PROD EN-US PV", false},
            {"https://www.merckmanuals.com/home/pages-with-widgets/figures?mode=list", "PROD EN-US CV", false},
            {"https://www.msdmanuals.com/pt/profissional/pages-with-widgets/figuras?mode=list", "PROD PT PV", false},
            {"https://www.msdmanuals.com/pt/casa/pages-with-widgets/figuras?mode=list", "PROD PT CV", false},
            {"https://www.msdmanuals.com/ja-jp/professional/pages-with-widgets/%E5%9B%B3?mode=list", "PROD JA PV", false},
            {"https://www.msdmanuals.com/ja-jp/home/pages-with-widgets/%E5%9B%B3?mode=list", "PROD JA CV", false},
            {"https://www.msdmanuals.com/fr/professional/pages-with-widgets/figures?mode=list", "PROD FR PV", false},
            {"https://www.msdmanuals.com/fr/accueil/pages-with-widgets/figures?mode=list", "PROD FR CV", false},
            {"https://www.msdmanuals.com/es/professional/pages-with-widgets/illustraciones?mode=list", "PROD ES PV", false},
            {"https://www.msdmanuals.com/es/hogar/pages-with-widgets/ilustraciones?mode=list", "PROD ES CV", false},
            {"https://www.msdmanuals.com/de/profi/pages-with-widgets/abbildungen?mode=list", "PROD DE PV", false},
            {"https://www.msdmanuals.com/de/heim/pages-with-widgets/abbildungen?mode=list", "PROD DE CV", false},
            {"https://www.msdmanuals.com/it/professionale/pages-with-widgets/figure?mode=list", "PROD IT PV", false},
            {"https://www.msdmanuals.com/it/casa/pages-with-widgets/figure?mode=list", "PROD IT CV", false},
            {"https://www.msdmanuals.com/ru/professional/pages-with-widgets/%D1%80%D0%B8%D1%81%D1%83%D0%BD%D0%BA%D0%B8?mode=list", "PROD RU PV", false},
            {"https://www.msdmanuals.com/ru/home/pages-with-widgets/%D1%80%D0%B8%D1%81%D1%83%D0%BD%D0%BA%D0%B8?mode=list", "PROD RU CV", true},
            {"https://www.msdmanuals.cn/professional/pages-with-widgets/figures?mode=list", "PROD CN PV", true},
            {"https://www.msdmanuals.cn/home/pages-with-widgets/figures?mode=list", "PROD CN CV", false},
            {"https://www.msdmanuals.com/ko/home/pages-with-widgets/%EA%B7%B8%EB%A6%BC?mode=list", "PROD KO CV", false},
            {"https://www.merckvetmanual.com/pages-with-widgets/figures?mode=list", "PROD MM VET", false},
            {"https://www.msdvetmanual.com/pages-with-widgets/figures?mode=list", "PROD MSD VET", false},
            {"https://www.msdmanuals.com/vi/professional/pages-with-widgets/c%C3%A1c-h%C3%ACnh-minh-h%E1%BB%8Da?mode=list", "PROD VI MSD PV", false},
            {"https://www.msdmanuals.com/uk/professional/pages-with-widgets/figures?mode=list", "PROD UK PV", false},
            {"https://www.msdmanuals.com/hi/home/pages-with-widgets/figures?mode=list", "PROD HI CV", false},
            {"https://www.msdmanuals.com/sw/home/pages-with-widgets/figures?mode=list", "PROD SW CV", false}
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
            verifyFigures();
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

    private void handlePopups() {
        try {
            System.out.println("Handling popups...");
            WebElement acceptCookies = wd.findElement(By.xpath("//*[@id='onetrust-accept-btn-handler']"));
            acceptCookies.click();
            System.out.println("Cookies popup handled.");
        } catch (Exception e) {
            System.out.println("Cookies popup not found.");
        }

        try {
            WebElement languageSelector = wd.findElement(By.xpath("//*[@class='ChineseModalPopup_languageSelectorPopupVersionButton__j7M_0']"));
            languageSelector.click();
            System.out.println("Language selector popup handled.");
        } catch (Exception e) {
            System.out.println("Language selector popup not found.");
        }
    }

    @AfterClass
    public void closeBrowser() throws Exception {
        getDate();
        System.out.println("================Browser Closed!================");
    }
}