import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class bigBasket {

	public static void main(String[] args) throws InterruptedException, ParseException, IOException {

		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\shivr\\Downloads\\chromedriver_win32\\chromedriver.exe");
		// Initialize browser
		WebDriver driver = new ChromeDriver();
		// ExcelDataPoolManager exclPool = new ExcelDataPoolManager();
		driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.get("https://www.bigbasket.com/");
		driver.manage().window().maximize();
		driver.findElement(By.xpath("//span[@class='hvc']")).click();
		driver.findElement(By.xpath("(//i[@class='caret pull-right'])[1]")).click();
		;
		WebElement city = driver.findElement(By.xpath("(//input[@type='search'])[1]"));
		ArrayList<WebElement> nameInList = new ArrayList<WebElement>();
		nameInList.add(city);
		city.click();
		city.clear();
		city.sendKeys("Hyderabad");
		city.sendKeys(Keys.ENTER);
		WebElement area = driver.findElement(By.xpath("(//input[@qa='areaInput'])[1]"));
		area.click();
		area.clear();
		area.sendKeys("Jub");
		Thread.sleep(1000);
		area.sendKeys(Keys.ENTER);
		driver.findElement(By.xpath("//button[@name='continue']")).click();
		Thread.sleep(3000);
		driver.findElement(By.xpath("//a[@ng-mouseover='vm.readyToShow = true']")).click();
		;
		driver.findElement(By.xpath("(//a[contains(text(),'Beverages')])[2]")).click();

		for (int i = 1; i < 3; i++) {

			String var = "(//button[@class='btn btn-default dropdown-toggle form-control'])[" + i + "]";
			Thread.sleep(8000);
			driver.findElement(By.xpath(var)).click();
			String varSecondTextElement = "((//button[@class='btn btn-default dropdown-toggle form-control'])[" + i
					+ "]//following-sibling::ul/li//following-sibling::li/a)[1]/span";
			String varSecondText = driver.findElement(By.xpath(varSecondTextElement)).getText();
			System.out.println(varSecondText);
			String spSecElement = "(((//button[@class='btn btn-default dropdown-toggle form-control'])[" + i
					+ "]//following-sibling::ul/li//following-sibling::li/a)[1]/span//following-sibling::span)[2]";
			String spSec = driver.findElement(By.xpath(spSecElement)).getText();
			System.out.println(spSec);
			String varSecond = "((//button[@class='btn btn-default dropdown-toggle form-control'])[" + i
					+ "]//following-sibling::ul/li//following-sibling::li/a)[1]";

			WebElement varSec = driver.findElement(By.xpath(varSecond));
			if (varSec.isDisplayed()) {
				driver.findElement(By.xpath(varSecond)).click();
				writeExcelByRow("C:\\Users\\shivr\\git\\MapprrWebAssignment\\MapprrWebAssignment\\src\\Maprrr.xls", "Sheet1", "variant", varSecondText, i);
				writeExcelByRow("C:\\Users\\shivr\\git\\MapprrWebAssignment\\MapprrWebAssignment\\src\\Maprrr.xls", "Sheet1", "selling_price", spSec, i);

			} else {
				String varFirst = "((//button[@class='btn btn-default dropdown-toggle form-control'])[1]//following-sibling::ul/li/a)[1]";
				driver.findElement(By.xpath(varFirst)).click();
			}
			String prodNamedynamicXpath = "(//h6[@ng-bind='vm.selectedProduct.p_brand'])[" + i + "]";
			String productName = driver.findElement(By.xpath(prodNamedynamicXpath)).getText();
			System.out.println(productName);
			writeExcelByRow("C:\\Users\\shivr\\git\\MapprrWebAssignment\\MapprrWebAssignment\\src\\Maprrr.xls", "Sheet1", "product_name", productName, i);

			String prodTypeDynamicXpath = "(//h6[@ng-bind='vm.selectedProduct.p_brand'])[" + i
					+ "]//following-sibling::a";
			String prodType = driver.findElement(By.xpath(prodTypeDynamicXpath)).getText();
			System.out.println(prodType);
			writeExcelByRow("C:\\Users\\shivr\\git\\MapprrWebAssignment\\MapprrWebAssignment\\src\\Maprrr.xls", "Sheet1", "product_type", prodType, i);
		}

		System.out.println("Testcase Passed");
		driver.close();

	}

	public static void writeExcelByRow(String XLS_FILE_PATH, String sSheetName, String columnName, String columnValue,
			int iRow) throws IOException {

		// ------------Declare Excel Sheet Variables----------------//
		InputStream minputStreamReadRow = null;
		OutputStream moutputStreamWriteRow = null;
		HSSFWorkbook mhssfwrkbokWorkbook;
		HSSFRow mhssfrowRow = null;
		HSSFCell mhssfcellCell = null;
		HSSFSheet msheetSheet;

		// ------------Declare Excel Sheet Variables----------------//
		minputStreamReadRow = new FileInputStream(XLS_FILE_PATH);
		moutputStreamWriteRow = null;

		mhssfwrkbokWorkbook = new HSSFWorkbook(minputStreamReadRow);
		mhssfrowRow = null;
		mhssfcellCell = null;
		msheetSheet = mhssfwrkbokWorkbook.getSheet(sSheetName);
		int colNo = -1;
		mhssfrowRow = msheetSheet.getRow(0);
		System.out.println("mhssfrowRow.getLastCellNum() " + mhssfrowRow.getLastCellNum());
		for (int i = 0; i < mhssfrowRow.getLastCellNum(); i++) {
			if (mhssfrowRow.getCell(i).getStringCellValue().equalsIgnoreCase(columnName)) {
				colNo = i;
				System.out.println("Column Number in Excel " + colNo);
			}
		}

		mhssfrowRow = msheetSheet.getRow(iRow);

		if (mhssfrowRow == null) {
			mhssfrowRow = msheetSheet.createRow(iRow);
		}

		mhssfcellCell = mhssfrowRow.getCell(colNo);
		if (mhssfcellCell == null) {
			mhssfcellCell = mhssfrowRow.createCell(colNo);
		}

		mhssfcellCell.setCellValue(columnValue);

		moutputStreamWriteRow = new FileOutputStream(XLS_FILE_PATH);
		mhssfwrkbokWorkbook.write(moutputStreamWriteRow);
		mhssfwrkbokWorkbook.close();

		if (moutputStreamWriteRow != null) {
			moutputStreamWriteRow.close();
			moutputStreamWriteRow = null;
		}
		moutputStreamWriteRow = null;

		if (minputStreamReadRow != null) {
			minputStreamReadRow.close();
			minputStreamReadRow = null;
		}
	}
}