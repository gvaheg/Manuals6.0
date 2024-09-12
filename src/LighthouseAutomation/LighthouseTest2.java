package LighthouseAutomation;

import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class LighthouseTest2 {

    // Run Lighthouse for a given URL with mobile emulation and multiple categories
    public static void runLighthouse(String url, String outputPath) {
        try {
            // Provide the correct path to the lighthouse executable
            String lighthousePath = "C:/Users/vahep/AppData/Roaming/npm/lighthouse.cmd"; // Adjust the path based on your installation

            // Build the command for mobile emulation and multiple categories
            List<String> command = new ArrayList<>();
            command.add(lighthousePath);
            command.add(url);
            command.add("--output");
            command.add("json");
            command.add("--output-path");
            command.add(outputPath);
            // Remove the --headless flag to make the browser visible
            command.add("--chrome-flags=--headless --user-agent='Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15A372 Safari/604.1' --window-size=375,812");
            command.add("--only-categories=performance,accessibility,best-practices,seo");  // Limit categories

            // Use ProcessBuilder for better process handling
            ProcessBuilder processBuilder = new ProcessBuilder(command);
            processBuilder.redirectErrorStream(true); // Merge stdout and stderr

            // Start the process
            Process process = processBuilder.start();

            // Capture the output stream
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            String line;
            while ((line = reader.readLine()) != null) {
                System.out.println(line);
            }

            // Wait for the process to finish
            int exitCode = process.waitFor();
            System.out.println("Lighthouse process for " + url + " exited with code: " + exitCode);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Parse the Lighthouse JSON report and extract required metrics and scores
    public static JsonObject parseLighthouseReport(String filePath) {
        JsonObject results = new JsonObject();
        try (FileReader reader = new FileReader(filePath)) {
            // Parse the JSON report
            JsonObject jsonObject = JsonParser.parseReader(reader).getAsJsonObject();

            // Extract required category scores and check if values are available
            results.addProperty("performance", getValue(jsonObject, "categories", "performance"));
            results.addProperty("accessibility", getValue(jsonObject, "categories", "accessibility"));
            results.addProperty("best-practices", getValue(jsonObject, "categories", "best-practices"));
            results.addProperty("seo", getValue(jsonObject, "categories", "seo"));

            // Extract specific performance metrics
            results.addProperty("first-contentful-paint", getAuditValue(jsonObject, "first-contentful-paint"));
            results.addProperty("largest-contentful-paint", getAuditValue(jsonObject, "largest-contentful-paint"));
            results.addProperty("total-blocking-time", getAuditValue(jsonObject, "total-blocking-time"));
            results.addProperty("cumulative-layout-shift", getAuditValue(jsonObject, "cumulative-layout-shift"));
            results.addProperty("speed-index", getAuditValue(jsonObject, "speed-index"));
        } catch (Exception e) {
            e.printStackTrace();
        }
        return results;
    }

    // Helper method to safely get category scores from the JSON report
    public static double getValue(JsonObject jsonObject, String category, String key) {
        try {
            return jsonObject.getAsJsonObject(category).getAsJsonObject(key).get("score").getAsDouble();
        } catch (Exception e) {
            return 0.0; // Default value if the score is not available
        }
    }

    // Helper method to safely get audit metrics from the JSON report
    public static String getAuditValue(JsonObject jsonObject, String key) {
        try {
            return jsonObject.getAsJsonObject("audits").getAsJsonObject(key).get("displayValue").getAsString();
        } catch (Exception e) {
            return "N/A"; // Default value if the metric is not available
        }
    }

    // Write results to Excel, including category scores and specific performance metrics
    public static void writeResultsToExcel(String fileName, List<String> urls, List<JsonObject> auditResults) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Lighthouse Results");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("URL");
        headerRow.createCell(1).setCellValue("Performance");
        headerRow.createCell(2).setCellValue("Accessibility");
        headerRow.createCell(3).setCellValue("Best Practices");
        headerRow.createCell(4).setCellValue("SEO");
        headerRow.createCell(5).setCellValue("First Contentful Paint");
        headerRow.createCell(6).setCellValue("Largest Contentful Paint");
        headerRow.createCell(7).setCellValue("Total Blocking Time");
        headerRow.createCell(8).setCellValue("Cumulative Layout Shift");
        headerRow.createCell(9).setCellValue("Speed Index");

        // Add the data rows
        for (int i = 0; i < urls.size(); i++) {
            Row row = sheet.createRow(i + 1);
            JsonObject result = auditResults.get(i);
            row.createCell(0).setCellValue(urls.get(i));
            row.createCell(1).setCellValue(result.get("performance").getAsDouble());
            row.createCell(2).setCellValue(result.get("accessibility").getAsDouble());
            row.createCell(3).setCellValue(result.get("best-practices").getAsDouble());
            row.createCell(4).setCellValue(result.get("seo").getAsDouble());
            row.createCell(5).setCellValue(result.get("first-contentful-paint").getAsString());
            row.createCell(6).setCellValue(result.get("largest-contentful-paint").getAsString());
            row.createCell(7).setCellValue(result.get("total-blocking-time").getAsString());
            row.createCell(8).setCellValue(result.get("cumulative-layout-shift").getAsString());
            row.createCell(9).setCellValue(result.get("speed-index").getAsString());
        }

        // Write to the Excel file
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
        }

        // Close the workbook
        workbook.close();
    }

    public static String sanitizeFileName(String url) {
        // Sanitize the URL to create a valid file name
        return url.replace("https://", "").replace("http://", "").replaceAll("[^a-zA-Z0-9.-]", "_");
    }

    public static void main(String[] args) {
        // List of URLs to audit
        List<String> urls = new ArrayList<>();
        urls.add("https://merckmanuals.com");
        urls.add("https://merckmanuals.com/home");
        urls.add("https://merckmanuals.com/professional");
        urls.add("https://merckmanuals.com/home/skin-disorders/acne-and-related-disorders/acne?query=acne");
        urls.add("https://merckmanuals.com/professional/eye-disorders/eyelid-and-lacrimal-disorders/dacryocystitis");
        urls.add("https://www.merckmanuals.com/professional/pages-with-widgets/figures?mode=list");

        // List to store audit results for each URL
        List<JsonObject> auditResults = new ArrayList<>();

        // Loop through each URL, run Lighthouse, parse the result, and store the audit results
        for (String url : urls) {
            String outputPath = "lighthouse-report-" + sanitizeFileName(url) + ".json";
            
            // Run Lighthouse for the current URL
            runLighthouse(url, outputPath);

            // Parse the audit results from the report
            JsonObject result = parseLighthouseReport(outputPath);
            auditResults.add(result);
        }

        // Write the results to an Excel file
        try {
            writeResultsToExcel("LighthouseResults.xlsx", urls, auditResults);
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Lighthouse audits completed and results written to LighthouseResults.xlsx");
    }
}
