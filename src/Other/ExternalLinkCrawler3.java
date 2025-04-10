package Other;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
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

public class ExternalLinkCrawler3 {

    private static Set<String> visitedUrls = new HashSet<>();
    private static Queue<String> urlQueue = new LinkedList<>();
    private static String baseUrl = "https://www.msdmanuals.com";
    private static String sitemapUrl = "https://www.msdmanuals.com/sitemap.ashx";
    private static WebDriver driver;
    private static Workbook workbook = new XSSFWorkbook();
    private static Sheet sheet = workbook.createSheet("ExternalURLs");
    private static int rowIndex = 1;
    private static int urlCounter = 1;

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
            System.out.println("Total URLs found: " + sitemapUrls.size());
            urlQueue.addAll(sitemapUrls);

            crawl();

            saveResultsToExcel();
        } finally {
            driver.quit();
            saveResultsToExcel();
        }
    }

    // Fetch URLs from sitemap
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

    // Crawl URLs
    public static void crawl() {
        int totalUrls = urlQueue.size();
        while (!urlQueue.isEmpty()) {
            String currentUrl = urlQueue.poll();
            if (visitedUrls.contains(currentUrl)) continue;

            try {
                Thread.sleep(500); // Slight delay to reduce server load
                driver.get(currentUrl);
                visitedUrls.add(currentUrl);
                System.out.println("[" + urlCounter + "/" + totalUrls + "] Crawling: " + currentUrl);
                urlCounter++;

                List<WebElement> links = getElementsWithRetry(By.xpath("//a[@href]"));
                for (WebElement link : links) {
                    try {
                        String href = link.getAttribute("href");
                        if (href == null || href.isEmpty() || !isValidUrl(href)) continue;

                        href = normalizeUrl(href);

                        if (isIgnoredUrl(href)) {
                            System.out.println("Ignored URL: " + href);
                            continue;
                        }

                        if (isExternalLink(href)) {
                            int statusCode = checkUrlStatus(href);
                            String merckWordPath = findMerckWordPath(href);
                            System.out.println("External URL: " + href + " - Status: " + statusCode);
                            writeToExcel(currentUrl, href, statusCode, merckWordPath);
                            saveResultsToExcel();
                        }
                    } catch (StaleElementReferenceException e) {
                        System.err.println("Stale element exception - Skipping link: " + e.getMessage());
                    }
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            } catch (Exception e) {
                System.err.println("Error processing URL: " + currentUrl + " - " + e.getMessage());
            }
        }
    }

    // Helper Methods
    public static List<WebElement> getElementsWithRetry(By locator) {
        int retries = 5;
        for (int i = 0; i < retries; i++) {
            try {
                return driver.findElements(locator);
            } catch (StaleElementReferenceException e) {
                System.err.println("Stale element - Retrying (" + (i + 1) + "/" + retries + ")");
            }
        }
        return Collections.emptyList();
    }

    public static boolean isIgnoredUrl(String url) {
        return ignoredUrls.stream().anyMatch(url::startsWith);
    }

    public static boolean isExternalLink(String url) {
        try {
            return !new URL(url).getHost().contains("msdmanuals.com");
        } catch (MalformedURLException e) {
            return false;
        }
    }

    public static String findMerckWordPath(String urlString) {
        try {
            driver.get(urlString);
            List<WebElement> elements = driver.findElements(By.xpath("//*[contains(text(), 'merck')]"));
            return elements.isEmpty() ? "OK" : elements.get(0).getTagName();
        } catch (Exception e) {
            return "Error";
        }
    }

    public static String normalizeUrl(String url) {
        return url.replaceAll("/+$", "").toLowerCase();
    }

    public static boolean isValidUrl(String url) {
        try {
            new URL(url).toURI();
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    public static int checkUrlStatus(String urlString) {
        try {
            HttpURLConnection connection = (HttpURLConnection) new URL(urlString).openConnection();
            connection.setRequestMethod("HEAD");
            connection.connect();
            return connection.getResponseCode();
        } catch (IOException e) {
            return -1;
        }
    }

    public static void writeToExcel(String sourceUrl, String externalUrl, int statusCode, String merckWordPath) {
        Row row = sheet.createRow(rowIndex++);
        row.createCell(0).setCellValue(sourceUrl);
        row.createCell(1).setCellValue(externalUrl);
        row.createCell(2).setCellValue(statusCode);
        row.createCell(3).setCellValue(merckWordPath);
    }

    public static void saveResultsToExcel() {
        try (FileOutputStream fileOut = new FileOutputStream("C:\\TestResults\\ExternalURLs.xlsx")) {
            workbook.write(fileOut);
            System.out.println("Results saved.");
        } catch (IOException e) {
            System.err.println("Error saving Excel: " + e.getMessage());
        }
    }
}
