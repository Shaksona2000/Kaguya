package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.io.FileOutputStream;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.util.*;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {

    public static WebDriver driver;

    public static void main(String[] args) {
        driver = new ChromeDriver();
        driver.manage().window().maximize();

        shakostanvartBichebi();
        // Close the browser
//        driver.quit();

    }
    public static void shakostanvartBichebi(){
        Map<String, String> URLSS;
        URLSS = new HashMap<String, String>();
        URLSS.put("silamaze da tavis movla extra j","https://extra.ge/catalog/silamaze-da-tavis-movla/344");



        String filePath = "C:\\Users\\sh.beridze\\Desktop\\extraaa.xlsx";
        Path path = Paths.get(filePath);

        // Create a Workbook
        Workbook workbook = new XSSFWorkbook();

        // Check if the file exists; if not, create a new one
        if (!Files.exists(path)) {
            try {
                 Files.createFile(path);
                System.out.println("File created: " + filePath);
            } catch (IOException e) {
                System.err.println("Failed to create file: " + filePath);
                e.printStackTrace();
            }
        } else {
            System.out.println("File already exists: " + filePath);
        }
        try {
            int sheetId = 1;
            for (Map.Entry<String, String> l : URLSS.entrySet()) {

                driver.get(l.getValue());

                // find the size of pages
                WebElement scrollToPages = driver.findElement(By.xpath("//ul[@class='_x_list-unstyled _x_mt-10 _x_flex _x_items-center _x_justify-center lg:_x_mt-25']"));
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("arguments[0].scrollIntoView(true);", scrollToPages);
                List<WebElement> pagerButtons = scrollToPages.findElements(By.xpath("//li[@class='_x_ml-5 _x_flex _x_flex-col']"));
                WebElement numberofPages = pagerButtons.get(pagerButtons.size() - 1);
                String quantityOfPages = numberofPages.getText();
                int numberOfButtons = Integer.parseInt(quantityOfPages);
                System.out.println(numberOfButtons);

                Sheet sheet = workbook.createSheet(l.getKey());
                // Create a Row
                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("Product");
                headerRow.createCell(1).setCellValue("ID");
                headerRow.createCell(2).setCellValue("Original price");
                headerRow.createCell(3).setCellValue("Saled price");
                headerRow.createCell(4).setCellValue("Discount Percent");

                int k = 1;
                for (int j = 1; j <= numberOfButtons; j++) {
                    String splittedUrl = l.getValue().split("=")[0];
                    driver.get(splittedUrl + "=" + j);

                    WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

                    WebElement mobileCategory = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@class='_x_grid _x_w-full _x_grid-cols-1 _x_justify-items-center _x_px-8 sm:_x_grid-cols-3 md:_x_justify-items-end md:_x_px-0 rg:_x_grid-cols-4 lg:_x_grid-cols-4']")));
                    js.executeScript("arguments[0].scrollIntoView(true);", mobileCategory);

                    List<WebElement> childElements = mobileCategory.findElements(By.xpath("//*[@class='_x_relative _x_flex _x_max-h-289 _x_w-full _x_flex-row _x_justify-between _x_border-b _x_border-dark-100 _x_pb-0 _x_pb-7 _x_pt-7 _x_text-black hover:_x_text-black sm:_x_flex-col sm:_x_p-5 sm:_x_py-17 md:_x_px-10 md:_x_py-9']"));
                    List<String> urls = new ArrayList<>();
                    Set<String> visitedUrls = new HashSet<>();

                    for (WebElement e : childElements) {
                        try {
                            WebElement aTag = wait.until(ExpectedConditions.elementToBeClickable(e.findElement(By.tagName("a"))));
                            String url = aTag.getAttribute("href");
                            urls.add(url);
                        } catch (org.openqa.selenium.NoSuchElementException ex) {
                            // Handle case where <a> tag is not found
                            System.err.println("No <a> tag found in element: " + e);
                            ex.printStackTrace();
                        }
                    }


                    for (String x : urls) {
                        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(1));
                        if (!visitedUrls.contains(x)) {
                            visitedUrls.add(x);
                            driver.get(x);
                           // WebElement price = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//div[@class='_x_mr-4 _x_flex _x_items-center'])[2]")));
                                //აქანე გავიჭედე არ მიკეთებს თრაი ქეჩს
                            WebElement price;
                            try {
                                // First, try to find the sale price element
                                price = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//div[@class='_x_mr-4 _x_flex _x_items-center'])[2]")));
                            } catch (TimeoutException e) {
                                // If the sale price element is not found, fall back to the regular price
                                price = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='_x_mr-4 _x_flex _x_items-center']")));
                            }


                            WebElement name = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//h1[@class='_x_m-none _x_break-words _x_font-bold _x_text-5 _x_text-dark md:_x_text-10'])[2]")));
                            WebElement ID = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//div[@class='_x_mb-3 _x_hidden md:_x_inline-block'])[2]")));

                            String priceText = (String) ((JavascriptExecutor) driver).executeScript("return arguments[0].innerText;", price);
                            String nameText = (String) ((JavascriptExecutor) driver).executeScript("return arguments[0].innerText;", name);
                            String phoneID = (String) ((JavascriptExecutor) driver).executeScript("return arguments[0].innerText;", ID);

                            String[] prices = priceText.split("\n");

                            System.out.println("\n");
                            System.out.println("=================================================\n");
                            System.out.println("#" + k + " - " + nameText + "\n");
                            System.out.println("Product " + phoneID + "\n");

                            Row row = sheet.createRow(k);
                            row.createCell(0).setCellValue(nameText);
                            row.createCell(1).setCellValue(phoneID);
                            if (prices.length > 1) {
                                int i = 1;
                                for (String s : prices) {
                                    if (i == 1) {

                                        row.createCell(3).setCellValue(s);
                                        System.out.println("saled Price- " + s);
                                    }
                                    if (i == 2) {
                                        row.createCell(2).setCellValue(s);
                                        System.out.println("Original Price- " + s);
                                    }

                                    if (i == 3) {
                                        row.createCell(4).setCellValue(s);
                                        System.out.println("Discount percent- " + s);
                                    }
                                    i++;
                                }
                            } else {
                                row.createCell(2).setCellValue(priceText.trim());
                                System.out.println("Original Price- " + priceText.trim());
                            }
                            System.out.println("=================================================\n");
                            k++;
                        }
                    }
                }
                sheetId++;
            }
        }catch (Exception ex){
            throw ex;
        }
        finally {
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
                System.out.println("Data written to file: " + filePath);
            } catch (IOException e) {
                e.printStackTrace();
            }

            // Close the workbook
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
