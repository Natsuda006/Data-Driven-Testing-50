import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import java.io.FileInputStream;
import java.io.IOException;

public class TestRegister {
    @Test
    void test01() throws IOException {
        System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");

        String path = "./excel/testdata.xlsx";
        FileInputStream fs = new FileInputStream(path);

        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int row = sheet.getLastRowNum() + 1;

        for(int i=1; i<row-1; i++) {
            WebDriver driver = new ChromeDriver();
            driver.get("http://localhost/sc_shortcourses/signup");

//            Row rows = sheet.getRow(i);
//            Cell cell = rows.createCell(4);
            XSSFRow currentRow = sheet.getRow(i);
            if (currentRow == null) {
                System.out.println("Skipping empty row: " + i);
                driver.quit();
                break;
            }

            //Thai
            String nameTitleTha = sheet.getRow(i).getCell(1).toString();
            new Select(driver.findElement(By.id("nameTitleTha"))).selectByValue(nameTitleTha);

            String firstnameTha = sheet.getRow(i).getCell(2).toString();
            driver.findElement(By.id("firstnameTha")).sendKeys(firstnameTha);

            String lastnameTha = sheet.getRow(i).getCell(3).toString();
            driver.findElement(By.id("lastnameTha")).sendKeys(lastnameTha);

            //Eng
            String nameTitleEng = sheet.getRow(i).getCell(4).toString();
            new Select(driver.findElement(By.id("nameTitleEng"))).selectByValue(nameTitleEng);

            String firstnameEng = sheet.getRow(i).getCell(5).toString();
            driver.findElement(By.id("firstnameEng")).sendKeys(firstnameEng);

            String lastnameEng = sheet.getRow(i).getCell(6).toString();
            driver.findElement(By.id("lastnameEng")).sendKeys(lastnameEng);

            String birthDate = sheet.getRow(i).getCell(7).toString();
            new Select(driver.findElement(By.id("birthDate"))).selectByValue(birthDate);

            String birthMonth = sheet.getRow(i).getCell(8).toString();
            new Select(driver.findElement(By.id("birthMonth"))).selectByValue(birthMonth);

            String birthYear = sheet.getRow(i).getCell(9).toString();
            new Select(driver.findElement(By.id("birthYear"))).selectByValue(birthYear);

            String idCard = sheet.getRow(i).getCell(10).toString();
            driver.findElement(By.id("idCard")).sendKeys(idCard);

            String password = sheet.getRow(i).getCell(11).toString();
            driver.findElement(By.id("password")).sendKeys(password);

            String mobile = sheet.getRow(i).getCell(12).toString();
            driver.findElement(By.id("mobile")).sendKeys(mobile);

            String email = sheet.getRow(i).getCell(13).toString();
            driver.findElement(By.id("email")).sendKeys(email);

            String address = sheet.getRow(i).getCell(14).toString();
            driver.findElement(By.id("address")).sendKeys(address);

            String province = sheet.getRow(i).getCell(15).toString();
            new Select(driver.findElement(By.id("province"))).selectByValue(province );

            String district = sheet.getRow(i).getCell(16).toString();
            driver.findElement(By.id("district")).sendKeys(district);

            String subDistrict = sheet.getRow(i).getCell(17).toString();
            driver.findElement(By.id("subDistrict")).sendKeys(subDistrict);

            String  postalCode = sheet.getRow(i).getCell(18).toString();
            driver.findElement(By.id("postalCode")).sendKeys(postalCode);

            String  acceptExcel = sheet.getRow(i).getCell(19).toString();
            if (acceptExcel.toLowerCase().equals("true")) {
                WebElement accept = driver.findElement(By.id("accept"));

                JavascriptExecutor js = (JavascriptExecutor) driver;
                if(!accept.isSelected()){
                    js.executeScript("arguments[0].click();", accept);
                }
            }
            WebElement submitBtn = driver.findElement(By.xpath("/html/body/section/div/div/form/div[6]/button"));
            submitBtn.submit();

            WebElement form = driver.findElement(By.xpath("/html/body/section/div/div/form"));
            form.submit();
            System.out.println("Register Successfully!");

            driver.quit();
        }
    }
}
