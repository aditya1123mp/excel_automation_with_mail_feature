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
import org.openqa.selenium.Keys as Keys
import org.apache.commons.lang.RandomStringUtils as RandomStringUtils
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import java.io.File as File
import java.io.FileInputStream as FileInputStream
import java.io.FileOutputStream as FileOutputStream
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

WebUI.callTestCase(findTestCase('LOGIN_TEST_CASE_FOR_ORANGE_HRM'), [:], FailureHandling.STOP_ON_FAILURE)

// Define the Excel file path and sheet
//String excelFilePath = "C:\\Users\\DELLOneDrive\\Desktop\\mail\\add_employee_data.xlsx"
String excelFilePath = 'C:\\Users\\DELL\\OneDrive\\Desktop\\mail\\add_employee_data.xlsx'

String sheetName = 'Sheet1'

// Create the Excel file if it doesn't exist
File file = new File(excelFilePath)

Workbook workbook

Sheet sheet

if (!(file.exists())) {
    workbook = new XSSFWorkbook()

    sheet = workbook.createSheet(sheetName)

    FileOutputStream fos = new FileOutputStream(file)

    workbook.write(fos)

    fos.close()

    workbook.close()
}

// Now open the existing file
FileInputStream fis = new FileInputStream(file)

workbook = new XSSFWorkbook(fis)

sheet = workbook.getSheet(sheetName)

if (sheet == null) {
    sheet = workbook.createSheet(sheetName)
}

// Add headers if the sheet is empty
if (sheet.getPhysicalNumberOfRows() == 0) {
    Row headerRow = sheet.createRow(0)

    headerRow.createCell(0).setCellValue('First_name')
	
	headerRow.createCell(1).setCellValue('Middle_name')

    headerRow.createCell(2).setCellValue('Last_name')

    headerRow.createCell(3).setCellValue('Username')

    headerRow.createCell(4).setCellValue('Password')
	
	headerRow.createCell(5).setCellValue('Employee_id')
}

// Generate random loop count
//int loopCount = Integer.parseInt(RandomStringUtils.randomNumeric(1)) + 1 // Adjust range as needed
int loopCount = (Integer.parseInt(RandomStringUtils.randomNumeric(2)) % 11) + 15

for (int i = 1; i <= loopCount; i++) {
    // Your existing script to add a new employee
    WebUI.callTestCase(findTestCase('ADDING_A_NEW_EMPLOYEE'), [:], FailureHandling.STOP_ON_FAILURE)

    // (For brevity, the employee addition steps are skipped here. Please include them from your script.)
    // Example for setting a username and password
    //String username = "User_" + RandomStringUtils.randomAlphabetic(6)
    //String password = "Pass_" + RandomStringUtils.randomNumeric(4)
    // Simulate setting these values in your application
    //GlobalVariable.User_name = username
    //GlobalVariable.Password = password
    // Store the username and password in Excel
    int rowIndex = sheet.getLastRowNum() + 1

    Row row = sheet.createRow(rowIndex)

    row.createCell(0).setCellValue(GlobalVariable.First_name)
	
	row.createCell(1).setCellValue(GlobalVariable.Middle_name)

    row.createCell(2).setCellValue(GlobalVariable.Last_name)

    row.createCell(3).setCellValue(GlobalVariable.User_name)

    row.createCell(4).setCellValue(GlobalVariable.Password) // Now proceed with the rest of your test case steps, such as verifying the employee has been added, etc.
	
	row.createCell(5).setCellValue(GlobalVariable.employee_id)
}

// Save the Excel file
fis.close()

FileOutputStream fos = new FileOutputStream(file)

workbook.write(fos)

fos.close()

workbook.close()

println('Employee data has been successfully saved in Excel.')

WebUI.callTestCase(findTestCase('logout_testcase'), [:], FailureHandling.STOP_ON_FAILURE)

WebUI.callTestCase(findTestCase('test_case_for_excel_automation_and_verifiying_data_in_excel'), [:], FailureHandling.STOP_ON_FAILURE)

