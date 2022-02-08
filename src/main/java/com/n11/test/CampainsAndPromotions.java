
package com.n11.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;


public class CampainsAndPromotions {
   static final String url      = "https://www.n11.com/";
   static WebDriver    driver   = null;
   static int          rowCount = 0;

   @SuppressWarnings("unlikely-arg-type")
   public static void main(String[] args) throws InterruptedException, IOException {
      System.setProperty("webdriver.chrome.driver", "src/resources/chromedriver.exe");
      driver = new ChromeDriver();
      driver.manage().window().maximize();

      driver.get(url);
      WebElement campains = driver.findElement(By.xpath("//a[text()='Kampanyalar']"));
      campains.click();

      WebElement campPromTab = driver.findElement(By.className("campPromTab"));
      WebElement categories = null;

      XSSFWorkbook workbook = new XSSFWorkbook();
      XSSFSheet spreadsheet = workbook.createSheet("Promotions List");
      XSSFRow row;
      Map<Integer, Object[]> camps = new TreeMap<Integer, Object[]>();
      int rowid = 0;
      List<WebElement> links = campPromTab.findElements(By.tagName("li"));
      for (int i = 1; i < links.size(); i++) {
         categories = driver.findElement(By.xpath("//span[text()='" + links.get(i).getText() + "']"));
         categories.click();
         System.out.println("*****");
         System.out.println(links.get(i).getText());
         System.out.println("*****");
         List<WebElement> promotionLinks = driver.findElement(By.xpath("//*[@id=\"contentCampaignPromotion\"]/div/div[2]/div/div[2]/div/section[" + i + "]"))
                                                 .findElements(By.tagName("a"));

         for (int j = 0; j < promotionLinks.size(); j++) {
            System.out.println(promotionLinks.get(j).getAttribute("href"));
            camps.put(rowCount, new Object[]{promotionLinks.get(j).getAttribute("href")});
            System.out.println(j);

            rowCount++;

         }
         Set<Integer> keyid = camps.keySet();

         for (Integer key : keyid) {
            row = spreadsheet.createRow(rowid);
            Object[] objectArr = camps.get(key);
            int cellid = 0;
            rowid++;
            for (Object obj : objectArr) {
               Cell cell = row.createCell(cellid++);
               cell.setCellValue((String) obj);
            }
         }
         Thread.sleep(1111);
      }
      FileOutputStream out = new FileOutputStream(new File("D:/n11.xlsx"));

      workbook.write(out);
      out.close();
      System.out.println("n11.xlsx written successfully");
   }
}
