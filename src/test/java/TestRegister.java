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

public class TestRegister {
    @Test
    void test101() throws IOException {
        System.setProperty("webdriver.chrome.driver", "./chromedriver/chromedriver.exe");

        // เปิดไฟล์ Excel
        String path = "./Excel/Sci.xlsx";
        FileInputStream fs = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int row = sheet.getLastRowNum() + 1;

        DataFormatter formatter = new DataFormatter(); // ใช้เพื่ออ่านค่าจากทุกประเภทของ Cell

        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));

        for (int i = 1; i < row - 1; i++) {
            driver.get("https://sc.npru.ac.th/sc_shortcourses/signup");

            // อ่านค่าจาก Excel แบบไม่สนใจว่าเป็น String หรือ Numeric
            String titleTha = formatter.formatCellValue(sheet.getRow(i).getCell(1));
            String firstNameTha = formatter.formatCellValue(sheet.getRow(i).getCell(2));
            String lastNameTha = formatter.formatCellValue(sheet.getRow(i).getCell(3));
            String titleEng = formatter.formatCellValue(sheet.getRow(i).getCell(4));
            String firstNameEng = formatter.formatCellValue(sheet.getRow(i).getCell(5));
            String lastNameEng = formatter.formatCellValue(sheet.getRow(i).getCell(6));
            String birthDate = formatter.formatCellValue(sheet.getRow(i).getCell(7));
            String birthMonth = formatter.formatCellValue(sheet.getRow(i).getCell(8));
            String birthYear = formatter.formatCellValue(sheet.getRow(i).getCell(9));
            String idCard = formatter.formatCellValue(sheet.getRow(i).getCell(10));
            String password = formatter.formatCellValue(sheet.getRow(i).getCell(11));
            String mobile = formatter.formatCellValue(sheet.getRow(i).getCell(12));
            String email = formatter.formatCellValue(sheet.getRow(i).getCell(13));
            String address = formatter.formatCellValue(sheet.getRow(i).getCell(14));
            String province = formatter.formatCellValue(sheet.getRow(i).getCell(15));
            String district = formatter.formatCellValue(sheet.getRow(i).getCell(16));
            String subDistrict = formatter.formatCellValue(sheet.getRow(i).getCell(17));
            String postalCode = formatter.formatCellValue(sheet.getRow(i).getCell(18));
            String acceptExcel = formatter.formatCellValue(sheet.getRow(i).getCell(19));

            // กรอกข้อมูล
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

            // ตรวจสอบว่าต้องคลิก checkbox หรือไม่
            WebElement accept = driver.findElement(By.id("accept"));
            JavascriptExecutor js = (JavascriptExecutor) driver;
            if (!accept.isSelected()){
                js.executeScript("arguments[0].click();",accept);
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

