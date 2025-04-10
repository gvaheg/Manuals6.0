package Other;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExternalLinkCrawlerFINAL1 {

    private static Set<String> visitedUrls = new HashSet<>();
    private static Queue<String> urlQueue = new LinkedList<>();
    private static String baseUrl = "https://www.msdmanuals.com";
    private static String sitemapUrl = "https://www.msdmanuals.com/sitemap.ashx";
    private static WebDriver driver;
    private static Workbook workbook = new XSSFWorkbook();
    private static Sheet sheet = workbook.createSheet("ExternalURLs");
    private static int rowIndex = 1;

    private static Set<String> ignoredUrls = new HashSet<>(Arrays.asList(
        "https://www.msdvetmanual.com",
        "https://www.msdprivacy.com",
        "https://www.facebook.com",
        "https://twitter.com",
        "https://www.essentialaccessibility.com",
        "mailto:msdmanualspermissions@msd.com",
        "https://play.google.com",
        "https://apps.apple.com",
        "https://www.msd.com"
    ));

    public static void main(String[] args) {
        System.setProperty("webdriver.gecko.driver", "C:\\SeleniumDrivers\\geckodriver.exe");
        FirefoxOptions options = new FirefoxOptions();
        options.setHeadless(true);
        driver = new FirefoxDriver(options);
        System.out.println("WebDriver initialized.");

        try {
            // Create header row in Excel
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Source URL");
            headerRow.createCell(1).setCellValue("External URL");
            headerRow.createCell(2).setCellValue("HTTP Status Code");
            headerRow.createCell(3).setCellValue("Merck Word Path");

            // Fetch URLs from sitemap
            List<String> sitemapUrls = fetchUrlsFromSitemap(sitemapUrl);
            urlQueue.addAll(sitemapUrls);

            crawl();

            saveResultsToExcel();
        } finally {
            driver.quit();
            saveResultsToExcel();
        }
    }
    public static int checkUrlStatus(String urlString) {
        try {
            URL url = new URL(urlString);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("HEAD");
            connection.setConnectTimeout(5000); // Timeout of 5 seconds
            connection.setReadTimeout(5000);    // Read timeout of 5 seconds
            connection.connect();
            return connection.getResponseCode();
        } catch (IOException e) {
            System.err.println("Error checking URL status: " + urlString + " - " + e.getMessage());
            return -1; // Return -1 in case of error
        }
    }

    public static List<String> fetchUrlsFromSitemap(String sitemapUrl) {
        List<String> urls = new ArrayList<>();
        try {
            URL url = new URL(sitemapUrl);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");

            BufferedReader in = new BufferedReader(new InputStreamReader(connection.getInputStream()));
            String inputLine;
            StringBuilder content = new StringBuilder();
            while ((inputLine = in.readLine()) != null) {
                content.append(inputLine);
            }
            in.close();

            Pattern pattern = Pattern.compile("<loc>(.*?)</loc>");
            Matcher matcher = pattern.matcher(content.toString());
            while (matcher.find()) {
                String foundUrl = matcher.group(1);
                urls.add(foundUrl);
                System.out.println("Sitemap URL: " + foundUrl);
            }
        } catch (IOException e) {
            System.err.println("Error fetching sitemap: " + e.getMessage());
        }
        return urls;
    }

    public static void crawl() {
        while (!urlQueue.isEmpty()) {
            String currentUrl = urlQueue.poll();

            if (visitedUrls.contains(currentUrl)) {
                continue;
            }

            try {
                // Slow down the crawl to prevent server blocking
                Thread.sleep(1000 + new Random().nextInt(1000));

                driver.get(currentUrl);
                visitedUrls.add(currentUrl);
                System.out.println("Crawling: " + currentUrl);

                List<WebElement> links = driver.findElements(By.xpath("//a[@href]"));

                for (WebElement link : links) {
                    String href = link.getAttribute("href");

                    if (href == null || href.isEmpty()) {
                        continue;
                    }

                    href = normalizeUrl(href);

                    if (isIgnoredUrl(href)) {
                        System.out.println("Ignored URL: " + href);
                        continue;
                    }

                    if (isExternalLink(href)) {
                        int statusCode = checkUrlStatus(href);
                        String merckWordPath = findMerckWordPath(href);
                        System.out.println("Source Page: " + currentUrl);
                        System.out.println("External URL: " + href);
                        System.out.println("HTTP Status Code: " + statusCode);
                        System.out.println("Merck Word Path: " + merckWordPath);
                        writeToExcel(currentUrl, href, statusCode, merckWordPath);
                        saveResultsToExcel();
                    }
                }
            } catch (Exception e) {
                System.err.println("Error processing URL: " + currentUrl + " - " + e.getMessage());
            }
        }
    }

    public static String normalizeUrl(String url) {
        while (url.endsWith("/")) {
            url = url.substring(0, url.length() - 1);
        }
        return url.toLowerCase();
    }

    public static boolean isExternalLink(String url) {
        try {
            URL linkUrl = new URL(url);
            return !linkUrl.getHost().contains("msdmanuals.com");
        } catch (MalformedURLException e) {
            return false;
        }
    }

    public static boolean isIgnoredUrl(String url) {
        for (String ignoredUrl : ignoredUrls) {
            if (url.equalsIgnoreCase(ignoredUrl) || url.startsWith(ignoredUrl + "/")) {
                return true;
            }
        }
        return false;
    }

    public static String findMerckWordPath(String urlString) {
        try {
            driver.get(urlString);
            List<WebElement> elements = driver.findElements(By.xpath("//*[contains(text(), 'merck')]"));
            if (elements.isEmpty()) {
                return "OK";
            }
            StringBuilder paths = new StringBuilder();
            for (WebElement element : elements) {
                paths.append(getElementPath(element)).append(";");
            }
            return paths.toString();
        } catch (Exception e) {
            return "Error";
        }
    }

    public static String getElementPath(WebElement element) {
        String path = element.getTagName();
        WebElement parent = element.findElement(By.xpath(".."));
        while (parent != null) {
            path = parent.getTagName() + " > " + path;
            try {
                parent = parent.findElement(By.xpath(".."));
            } catch (Exception e) {
                break;
            }
        }
        return path;
    }

    public static void writeToExcel(String sourceUrl, String externalUrl, int statusCode, String merckWordPath) {
        Row row = sheet.createRow(rowIndex++);
        row.createCell(0).setCellValue(sourceUrl);
        row.createCell(1).setCellValue(externalUrl);
        row.createCell(2).setCellValue(statusCode);
        row.createCell(3).setCellValue(merckWordPath);
    }

    public static void saveResultsToExcel() {
        String outputFilePath = "C:\\TestResults\\ExternalURLs.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
            workbook.write(fileOut);
            System.out.println("Results saved to: " + outputFilePath);
        } catch (IOException e) {
            System.err.println("Error saving Excel file: " + e.getMessage());
        }
    }
}
