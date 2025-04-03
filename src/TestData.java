import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class TestData {
    @Test
    void test101() throws IOException {
        System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");

        String path = "./Excel/testData.xlsx";
        FileInputStream fs = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int row = sheet.getLastRowNum() + 1;

        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));

        for (int i = 1; i < row - 1; i++) {
            driver.get("http://localhost/sc_shortcourses/signup");

            // อ่านค่าจาก Excel
            String titleTha = getCellValue(sheet.getRow(i), 1);
            String firstNameTha = getCellValue(sheet.getRow(i), 2);
            String lastNameTha = getCellValue(sheet.getRow(i), 3);
            String titleEng = getCellValue(sheet.getRow(i), 4);
            String firstNameEng = getCellValue(sheet.getRow(i), 5);
            String lastNameEng = getCellValue(sheet.getRow(i), 6);
            String birthDate = getCellValue(sheet.getRow(i), 7);
            String birthMonth = getCellValue(sheet.getRow(i), 8);
            String birthYear = getCellValue(sheet.getRow(i), 9);
            String idCard = getCellValue(sheet.getRow(i), 10);
            String password = getCellValue(sheet.getRow(i), 11);
            String mobile = getCellValue(sheet.getRow(i), 12);
            String email = getCellValue(sheet.getRow(i), 13);
            String address = getCellValue(sheet.getRow(i), 14);
            String province = getCellValue(sheet.getRow(i), 15);
            String district = getCellValue(sheet.getRow(i), 16);
            String subDistrict = getCellValue(sheet.getRow(i), 17);
            String postalCode = getCellValue(sheet.getRow(i), 18);

            new Select(driver.findElement(By.id("nameTitleTha"))).selectByVisibleText(titleTha);
            driver.findElement(By.id("firstnameTha")).sendKeys(firstNameTha);
            driver.findElement(By.id("lastnameTha")).sendKeys(lastNameTha);

            new Select(driver.findElement(By.id("nameTitleEng"))).selectByVisibleText(titleEng);
            driver.findElement(By.id("firstnameEng")).sendKeys(firstNameEng);
            driver.findElement(By.id("lastnameEng")).sendKeys(lastNameEng);

            driver.findElement(By.id("birthDate")).sendKeys(birthDate);
            driver.findElement(By.id("birthMonth")).sendKeys(birthMonth);
            driver.findElement(By.id("birthYear")).sendKeys(birthYear);
            driver.findElement(By.id("idCard")).sendKeys(idCard);
            driver.findElement(By.id("password")).sendKeys(password);
            driver.findElement(By.id("mobile")).sendKeys(mobile);
            driver.findElement(By.id("email")).sendKeys(email);
            driver.findElement(By.id("address")).sendKeys(address);

            new Select(driver.findElement(By.id("province"))).selectByVisibleText(province);
            driver.findElement(By.id("district")).sendKeys(district);
            driver.findElement(By.id("subDistrict")).sendKeys(subDistrict);
            driver.findElement(By.id("postalCode")).sendKeys(postalCode);

            WebElement accept = driver.findElement(By.id("accept"));
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", accept);
            if (!accept.isSelected()) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", accept);
            }

            WebElement submitButton = driver.findElement(By.xpath("/html/body/section/div/div/form/div[6]/button"));
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", submitButton);


            WebElement alertTitle = driver.findElement(By.id("swal2-title"));
            assertEquals("ลงทะเบียนสำเร็จ", alertTitle.getText());

        }

        driver.quit();
        workbook.close();
        fs.close();
    }

    private String getCellValue(Row row, int cellIndex) {
        if (row.getCell(cellIndex) != null) {
            return row.getCell(cellIndex).toString().trim();
        }
        return "";
    }
}