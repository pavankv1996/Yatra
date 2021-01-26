package Peerxp;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Basic {
String xlFilePath="C:\\Users\\Karthik Kedlaya\\eclipse-workspace\\Yatra.xlsx";
String SheetName="Sheet1";
ExcelApiTest eat = null;
WebDriver driver = null;

@BeforeClass
public void init() throws InterruptedException
{
	System.setProperty("webdriver.chrome.driver", "./driver/chromedriver.exe");
	driver = new ChromeDriver();
	driver.manage().window().maximize();
	Thread.sleep(1000);
	driver.get("www.yatra.com");
}

 @Test(dataProvider="UserData")
 public void filllog(String From, String To )throws Exception
 {
	 driver.findElement(By.xpath("//input[@name='flight_origin']")).sendKeys(From);
	 Thread.sleep(3000);
	 driver.findElement(By.xpath("//input[@name='flight_destination']")).sendKeys(To);
	 Thread.sleep(3000);
	 driver.findElement(By.xpath("//input[@id='BE_flight_flsearch_btn']")).click();

 }		      
 @DataProvider(name="UserData")
public Object[][] userlog() throws Exception
{
	Object[][] data = testData(xlFilePath,SheetName);
	return data;
}

public Object[][] testData(String xlFilePath, String SheetName) throws Exception
{
	Object[][] excelData = null;
	eat = new ExcelApiTest(xlFilePath);
	
	int rows = eat.getRowCount(SheetName);
	int columns = eat.getcolumnCount(SheetName);
	
	excelData = new Object [rows-1][columns];
	for(int i=1;i<rows;i++)
	{
		for(int j=0;j<columns;j++)
		{
			excelData[i-1][j]= eat.getcellData(SheetName,j,i);
		}
	}
	return excelData;
}			      
}  
	
