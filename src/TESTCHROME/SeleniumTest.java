package TESTCHROME;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class SeleniumTest {
   public static void main(String[] args) throws InterruptedException {
      // Set the path of the ChromeDriver executable
      System.setProperty("webdriver.chrome.driver", "C:\\SeleniumDrivers\\chromedriver.exe");

      // Create a new instance of the ChromeDriver
      WebDriver driver = new ChromeDriver();

      // Navigate to a web page
      driver.get("https://www.google.com");
      Thread.sleep(10000);
      // Perform some actions on the web page
      // ...

      // Close the browser
      //driver.quit();
   }
}
