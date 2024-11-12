package Other;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.StringReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

public class UrlStatusChecker {

    public static void main(String[] args) {
        System.out.println("Starting URL status check...");

        // Create Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("URL Status");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ID");
        headerRow.createCell(1).setCellValue("URL");
        headerRow.createCell(2).setCellValue("Status Code");
        headerRow.createCell(3).setCellValue("Status Message");

        // Specify output file path and ensure the directory exists
        String outputFilePath = "C:\\TestResults\\URL_Status_Report.xlsx";
        File outputFile = new File(outputFilePath);
        outputFile.getParentFile().mkdirs(); // Create directory if it doesn't exist

        int rowNum = 1; // Start from the second row since the first is the header
        int idCounter = 1; // Unique ID for each URL

        try {
            // Retrieve URLs from sitemap
            List<String> urls = getUrlsFromSitemap("https://www.msdmanuals.com/sitemap.ashx");
            System.out.println("URLs retrieved from sitemap: " + urls.size());

            // Loop through each URL and check its status code
            for (String url : urls) {
                System.out.println("Checking URL: " + url);
                int statusCode = -1;
                String statusMessage = "";

                try {
                    // Open a connection and get the status code
                    HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
                    connection.setRequestMethod("GET");
                    connection.setInstanceFollowRedirects(false); // Disable automatic redirects
                    connection.connect();

                    statusCode = connection.getResponseCode();
                    statusMessage = connection.getResponseMessage();
                    connection.disconnect();
                } catch (IOException e) {
                    statusMessage = "Connection Failed";
                    e.printStackTrace();
                }

                // Create a new row in the Excel sheet
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(idCounter++);
                row.createCell(1).setCellValue(url);
                row.createCell(2).setCellValue(statusCode);
                row.createCell(3).setCellValue(statusMessage);

                System.out.println("URL: " + url + " | Status Code: " + statusCode + " | Message: " + statusMessage);

                // Save Excel file after each URL is processed
                try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                    workbook.write(fileOut);
                }
            }

            System.out.println("Excel file 'URL_Status_Report.xlsx' created successfully at " + outputFilePath);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    // Method to fetch and parse URLs from the sitemap
    private static List<String> getUrlsFromSitemap(String sitemapUrl) {
        List<String> urls = new ArrayList<>();
        HttpURLConnection connection = null;
        try {
            connection = (HttpURLConnection) new URL(sitemapUrl).openConnection();
            connection.setRequestMethod("GET");
            connection.setRequestProperty("User-Agent", "Mozilla/5.0");

            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(new InputSource(new StringReader(new String(connection.getInputStream().readAllBytes()))));
            NodeList urlNodes = doc.getElementsByTagName("loc");

            for (int i = 0; i < urlNodes.getLength(); i++) {
                urls.add(urlNodes.item(i).getTextContent());
            }
            System.out.println("Sitemap URLs parsed successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }
        return urls;
    }
}
