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

public class Quizzes extends All_Functions {
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
/*
            {"https://www.merckmanuals.com/professional/pages-with-widgets/quizzes?mode=list", "PROD EN-US PV", false},
            {"https://www.merckmanuals.com/home/pages-with-widgets/quizzes?mode=list", "PROD EN-US CV", false},
            {"https://www.msdmanuals.com/pt/profissional/pages-with-widgets/testes?mode=list", "PROD PT PV", false},
            {"https://www.msdmanuals.com/pt/casa/pages-with-widgets/testes?mode=list", "PROD PT CV", false},
            {"https://www.msdmanuals.com/ja-jp/professional/pages-with-widgets/%E3%82%AF%E3%82%A4%E3%82%BA?mode=list", "PROD JA PV", false},
            {"https://www.msdmanuals.com/ja-jp/home/pages-with-widgets/%E3%82%AF%E3%82%A4%E3%82%BA?mode=list", "PROD JA CV", false},
            {"https://www.msdmanuals.com/fr/professional/pages-with-widgets/questionnaires?mode=list", "PROD FR PV", false},
            {"https://www.msdmanuals.com/fr/accueil/pages-with-widgets/questionnaires?mode=list", "PROD FR CV", false},
            {"https://www.msdmanuals.com/es/professional/pages-with-widgets/cuestionarios?mode=list", "PROD ES PV", false},
            {"https://www.msdmanuals.com/es/hogar/pages-with-widgets/cuestionarios?mode=list", "PROD ES CV", false},
            {"https://www.msdmanuals.com/de/profi/pages-with-widgets/quizfragen?mode=list", "PROD DE PV", false},
            {"https://www.msdmanuals.com/de/heim/pages-with-widgets/quizfragen?mode=list", "PROD DE CV", false},
            {"https://www.msdmanuals.com/it/professionale/pages-with-widgets/quizzes?mode=list", "PROD IT PV", false},
            {"https://www.msdmanuals.com/it/casa/pages-with-widgets/quiz?mode=list", "PROD IT CV", false},
            {"https://www.msdmanuals.com/ru/professional/pages-with-widgets/%D0%BE%D0%BF%D1%80%D0%BE%D1%81%D1%8B?mode=list", "PROD RU PV", false},
            {"https://www.msdmanuals.com/ru/home/pages-with-widgets/%D0%BE%D0%BF%D1%80%D0%BE%D1%81%D1%8B?mode=list", "PROD RU CV", true},
            {"https://www.msdmanuals.cn/professional/pages-with-widgets/quizzes?mode=list", "PROD CN PV", true},
            {"https://www.msdmanuals.cn/home/pages-with-widgets/quizzes?mode=list", "PROD CN CV", false},
            {"https://www.msdmanuals.com/ko/home/pages-with-widgets/%ED%80%B4%EC%A6%88?mode=list", "PROD KO CV", false},
            {"https://www.merckvetmanual.com/pages-with-widgets/quizzes?mode=list", "PROD MM VET", false},
            {"https://www.msdvetmanual.com/pages-with-widgets/quizzes?mode=list", "PROD MSD VET", false},
            */
            {"https://www.msdmanuals.com/professional/pages-with-widgets/quizzes?mode=list", "PROD EN MSD PV", false},
            {"https://www.msdmanuals.com/home/pages-with-widgets/quizzes?mode=list", "PROD EN MSD CV", false},
            {"https://www.msdmanuals.com/ar/home/pages-with-widgets/quizzes?mode=list", "PROD AR CV", false},
            {"https://www.msdmanuals.com/vi/professional/pages-with-widgets/quizzes?mode=list", "PROD VI PV", false},
            {"https://www.msdmanuals.com/hi/home/pages-with-widgets/quizzes?mode=list", "PROD HI PV", false}

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
            verifyQuizzes();
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