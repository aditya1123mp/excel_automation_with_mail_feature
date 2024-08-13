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



WebUI.delay(3)

WebUI.click(findTestObject('to_verify_data/myinfo_option_from_side_bar'))

First_name = WebUI.getAttribute(findTestObject('Object Repository/to_verify_data/first_name_text_box'), 'value')

System.out.println(First_name)

Middle_name = WebUI.getAttribute(findTestObject('Object Repository/to_verify_data/middle_name_text_box'), 'value')

System.out.println(Middle_name)

Last_name = WebUI.getAttribute(findTestObject('Object Repository/to_verify_data/last_name_text_box'), 'value')

System.out.println(Last_name)

employee_id = WebUI.getAttribute(findTestObject('Object Repository/to_verify_data/employee_id'), 'value')

System.out.println(employee_id)

if (GlobalVariable.First_name == First_name) {
	KeywordUtil.markPassed('first name is same')
} else {
	KeywordUtil.markFailed('first name is not same')
}


if (GlobalVariable.Middle_name == Middle_name) {
	KeywordUtil.markPassed('middle name is same')
} else {
	KeywordUtil.markFailed('middle name is not same')
}


if (GlobalVariable.Last_name == Last_name) {
	KeywordUtil.markPassed('last name is same')
} else {
	KeywordUtil.markFailed('last name is not same')
}


if (GlobalVariable.employee_id == employee_id) {
	KeywordUtil.markPassed('employee id is same')
} else {
	KeywordUtil.markFailed('employee id is not same')
}


