package Resources;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.time.Duration;

public class MSDManualsPopupHandler {

    public static void main(String[] args) {
        // Set the GeckoDriver path
        System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");

        // Initialize WebDriver
        WebDriver driver = new FirefoxDriver();

        try {
            // Navigate to the page
            driver.get("https://www.msdmanuals.com/professional/pages-with-widgets/figures?mode=list");

            // Maximize the window
            driver.manage().window().maximize();

            // Set explicit wait
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

            // Handle any potential cookies or consent popups
            try {
                WebElement consentButton = wait.until(ExpectedConditions.elementToBeClickable(By.id("onetrust-accept-btn-handler")));
                consentButton.click();
            } catch (Exception ignored) {
                // No consent popup present
            }

            // Total number of items to iterate through
            int totalItems = 320;
            System.out.println("Total number of list items: " + totalItems);

            // Iterate over each item using dynamic XPath
            for (int i = 1; i <= totalItems; i++) {
                try {
                    // Locate the <a> element dynamically
                    WebElement listItem = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(".//main/div[1]//div[1]//div/div[2]/ul/li[" + i + "]/a")));

                    // Scroll into view
                    ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", listItem);

                    // Click the link
                    ((JavascriptExecutor) driver).executeScript("arguments[0].click();", listItem);

                    // Wait for the popup to appear
                    WebElement popup = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(@class, 'PopupContent_multiMediaDetailsContainer')]")));

                    // Locate the title within the popup
                    WebElement titleElement = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(@class, 'PopupContent_multiMediaTitle')]")));
                    String popupTitle = titleElement.getText().trim();

                    // Check if the title is empty
                    if (popupTitle.isEmpty()) {
                        System.out.println("Popup title not found for this item.");
                    } else {
                        System.out.println("Popup title: " + popupTitle);
                    }

                    // Close the popup
                    WebElement closeButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(@class, 'PopupContent_closeIcon')]")));
                    closeButton.click();

                    // Add a short delay to ensure proper closure
                    Thread.sleep(1000);
                } catch (Exception e) {
                    System.out.println("Error processing item index: " + i);
                    e.printStackTrace();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // Quit the WebDriver
            driver.quit();
        }
    }
}