import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable
import org.apache.commons.lang.RandomStringUtils as RandomStringUtils
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import java.io.File as File
import java.io.FileInputStream as FileInputStream
import java.io.FileOutputStream as FileOutputStream
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject



// Define the Excel file path and sheet
String excelFilePath = 'C:\\Users\\DELL\\OneDrive\\Desktop\\mail\\add_employee_data.xlsx'

String sheetName = 'Sheet1'

// Open the Excel file
FileInputStream fis = new FileInputStream(new File(excelFilePath))

Workbook workbook = new XSSFWorkbook(fis)

Sheet sheet = workbook.getSheet(sheetName)

// Iterate through the rows and use the data to log in
for (int i = 1; i <= sheet.getLastRowNum(); i++) {
    // Start from 1 to skip the header row
	WebUI.delay(3)
	
    Row row = sheet.getRow(i)

    String firstName = row.getCell(0).getStringCellValue()

    String Middlename = row.getCell(1).getStringCellValue()

    String lastName = row.getCell(2).getStringCellValue()

    String username = row.getCell(3).getStringCellValue()

    String password = row.getCell(4).getStringCellValue()

    String Employee_id = row.getCell(5).getStringCellValue()

    // Print the data to verify (Optional)
    println((('Logging in with Username: ' + username) + ' and Password: ') + password)

    // Call the login test case and pass the username and password
    GlobalVariable.User_name = username

    GlobalVariable.Password = password // Add additional steps if needed, like verifying the login was successful, etc.

    GlobalVariable.First_name = firstName

    GlobalVariable.Last_name = lastName

    GlobalVariable.Middle_name = Middlename

    GlobalVariable.employee_id = Employee_id

    WebUI.callTestCase(findTestCase('test_case_for_new_login_creds_excel_sheet'), [:], FailureHandling.STOP_ON_FAILURE)

    WebUI.callTestCase(findTestCase('test_case_to_verify_data'), [:], FailureHandling.STOP_ON_FAILURE)
	
	WebUI.callTestCase(findTestCase('logout_testcase'), [:], FailureHandling.STOP_ON_FAILURE)
}

// Close the workbook and file input stream
fis.close()

workbook.close()

println('Login process completed using data from Excel.')



