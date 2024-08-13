import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import org.apache.commons.lang.RandomStringUtils as RandomStringUtils
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.util.KeywordUtil as KeywordUtil
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.ConditionType as ConditionType
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webui.driver.DriverFactory
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.WebDriver
import org.openqa.selenium.By
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.support.ui.WebDriverWait
import org.openqa.selenium.support.ui.ExpectedConditions
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import org.openqa.selenium.WebElement
import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory

// Wait for the page to load
WebUI.delay(3)

// Click on the upload area to make the file input visible
WebUI.click(findTestObject('to_upload_profile_picture/employee_image_wrapper'))
WebUI.delay(1)

// Click the upload button to open the file dialog
//WebUI.click(findTestObject('Object Repository/to_upload_profile_picture/employee_image_action'))

// Initialize WebDriver and WebDriverWait
WebDriver driver = DriverFactory.getWebDriver()
WebDriverWait wait = new WebDriverWait(driver, 10)

// Wait for the file input element to be present

	WebElement fileInput = wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector('input[type="file"]')))
	
	// Make the file input element visible if it's not
	((JavascriptExecutor) driver).executeScript("arguments[0].style.display = 'block';", fileInput)
	
	// Set the file path to the file input element
	String filePath = "C:\\Users\\DELL\\OneDrive\\Desktop\\profile_picture.jpg" // Change this to the path of your file
	fileInput.sendKeys(filePath)
	
	// Optionally, click the upload or submit button if needed
	WebUI.click(findTestObject('Object Repository/to_upload_profile_picture/save_button'))
	
	WebUI.delay(1)
	
	WebUI.click(findTestObject('to_verify_data/myinfo_option_from_side_bar'))
