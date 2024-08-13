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

WebUI.verifyElementPresent(findTestObject('Object Repository/LOGIN_TEST_CASE_FOR_ORANGE_HRM/username_text_box_login_page'),10)

WebUI.click(findTestObject('Object Repository/LOGIN_TEST_CASE_FOR_ORANGE_HRM/username_text_box_login_page'))

WebUI.setText(findTestObject('Object Repository/LOGIN_TEST_CASE_FOR_ORANGE_HRM/username_text_box_login_page'), GlobalVariable.extractedUsername)

WebUI.verifyElementPresent(findTestObject('Object Repository/LOGIN_TEST_CASE_FOR_ORANGE_HRM/password_text_box_for_login_page'),10)

WebUI.click(findTestObject('Object Repository/LOGIN_TEST_CASE_FOR_ORANGE_HRM/password_text_box_for_login_page'))

WebUI.setText(findTestObject('Object Repository/LOGIN_TEST_CASE_FOR_ORANGE_HRM/password_text_box_for_login_page'), GlobalVariable.extractedPassword)

WebUI.verifyElementPresent(findTestObject('Object Repository/LOGIN_TEST_CASE_FOR_ORANGE_HRM/login_button_for_login_page'),10)

WebUI.click(findTestObject('Object Repository/LOGIN_TEST_CASE_FOR_ORANGE_HRM/login_button_for_login_page'))



