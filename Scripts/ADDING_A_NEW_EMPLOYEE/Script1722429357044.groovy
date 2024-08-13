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
import com.kms.katalon.core.testobject.ConditionType


WebUI.verifyElementPresent(findTestObject('ADDING_A_NEW_EMPLOYEE/pim_option_from_side_bar'), 10)

WebUI.click(findTestObject('ADDING_A_NEW_EMPLOYEE/pim_option_from_side_bar'))

WebUI.waitForElementPresent(findTestObject('ADDING_A_NEW_EMPLOYEE/adding_new_employee_button_in_pim_page'), 10)

WebUI.click(findTestObject('ADDING_A_NEW_EMPLOYEE/adding_new_employee_button_in_pim_page'))

WebUI.verifyElementPresent(findTestObject('ADDING_A_NEW_EMPLOYEE/create_login_details_checkbox'), 10)

WebUI.click(findTestObject('ADDING_A_NEW_EMPLOYEE/create_login_details_checkbox'))

WebUI.verifyElementPresent(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_enabled_radio_button'), 
    10)

WebUI.click(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_enabled_radio_button'))


WebUI.verifyElementPresent(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_Employee_first_name'), 10)

WebUI.click(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_Employee_first_name'))

WebUI.setText(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_Employee_first_name'), 'First_name_' + RandomStringUtils.randomAlphabetic(
	7))


WebUI.verifyElementPresent(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_middle_name'), 10)

WebUI.click(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_middle_name'))

WebUI.setText(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_middle_name'), 'Middle_name_' + RandomStringUtils.randomAlphabetic(
	7))


WebUI.verifyElementPresent(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_last_name'), 10)

WebUI.click(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_last_name'))

WebUI.setText(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_last_name'), 'Last_name_' + RandomStringUtils.randomAlphabetic(
	7))


//WebUI.verifyElementPresent(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_id'), 10)

//WebUI.clearText(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_id'))

//WebUI.click(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_id'))

//WebUI.setText(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_id'), RandomStringUtils.randomNumeric(6))

// Generate the random string
String randomString_001 = RandomStringUtils.randomNumeric(6)
String newValue_001 = "333" + randomString_001

// Define the XPath
String xpath_for_employee_id = "//label[text()='Employee Id']/ancestor::div[@class='oxd-input-group__label-wrapper']/following-sibling::div//input[@class='oxd-input oxd-input--active']"

// Locate the input field using the XPath
TestObject inputField_001 = new TestObject().addProperty("xpath", ConditionType.EQUALS, xpath_for_employee_id)

// Use JavaScript to set the value and trigger the input event
WebUI.executeJavaScript(
	"""
    var element = document.evaluate("${xpath_for_employee_id}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (element) {
        element.value = '${newValue_001}';
        var event = new Event('input', { bubbles: true });
        element.dispatchEvent(event);
    }
    """, 
	null
)

WebUI.delay(3)

//WebUI.verifyElementPresent(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_username'), 10)

//WebUI.click(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_username'))

//WebUI.setText(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_username'), 'Fiboo' + RandomStringUtils.randomNumeric(4) + '@@')

// Generate the random string for user id
String randomString_for_password = RandomStringUtils.randomNumeric(6)
String new_Value_for_password = "Fiboo" + randomString_for_password + '@@'

// Define the XPath
String xpath_for_password = "//label[text()='Username']/ancestor::div[@class='oxd-input-group__label-wrapper']/following-sibling::div//input[@class='oxd-input oxd-input--active']"

// Locate the input field using the XPath
TestObject inputField_for_password = new TestObject().addProperty("xpath", ConditionType.EQUALS, xpath_for_password)


// Use JavaScript to set the value and trigger the input event
WebUI.executeJavaScript(
	"""
    var element = document.evaluate("${xpath_for_password}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (element) {
        element.value = '${new_Value_for_password}';
        var event = new Event('input', { bubbles: true });
        element.dispatchEvent(event);
    }
    """, 
	null
)

WebUI.delay(1)



GlobalVariable.Password = WebUI.getAttribute(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_username'),'value')

System.out.println(GlobalVariable.Password)


//WebUI.verifyElementPresent(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_password'), 10)

//WebUI.click(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_password'))

//WebUI.setText(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_password'),GlobalVariable.Password)

// Generate the random string for password
//String randomString_for_password_001 = RandomStringUtils.randomNumeric(4)
String new_Value_for_password_001 = GlobalVariable.Password

// Define the XPath
String xpath_for_password_001 = "//label[text()='Password']/ancestor::div[@class='oxd-input-group__label-wrapper']/following-sibling::div//input[@class='oxd-input oxd-input--active']"

// Locate the input field using the XPath
TestObject inputField_for_password_001 = new TestObject().addProperty("xpath", ConditionType.EQUALS, xpath_for_password_001)


// Use JavaScript to set the value and trigger the input event
WebUI.executeJavaScript(
	"""
    var element = document.evaluate("${xpath_for_password_001}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (element) {
        element.value = '${new_Value_for_password_001}';
        var event = new Event('input', { bubbles: true });
        element.dispatchEvent(event);
    }
    """, 
	null
)

WebUI.delay(1)


//WebUI.verifyElementPresent(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_confirm_password'), 10)
// Generate the random string for password
//String randomString_for_password_001 = RandomStringUtils.randomNumeric(4)
String new_Value_for_password_002 = GlobalVariable.Password

// Define the XPath
String xpath_for_password_002 = "//label[text()='Confirm Password']/ancestor::div[@class='oxd-input-group__label-wrapper']/following-sibling::div//input[@class='oxd-input oxd-input--active']"

// Locate the input field using the XPath
TestObject inputField_for_password_002 = new TestObject().addProperty("xpath", ConditionType.EQUALS, xpath_for_password_002)


// Use JavaScript to set the value and trigger the input event
WebUI.executeJavaScript(
	"""
    var element = document.evaluate("${xpath_for_password_002}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (element) {
        element.value = '${new_Value_for_password_002}';
        var event = new Event('input', { bubbles: true });
        element.dispatchEvent(event);
    }
    """, 
	null
)

WebUI.delay(1)

//WebUI.click(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_confirm_password'))

//WebUI.setText(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_confirm_password'),GlobalVariable.Password)



// Generate the random string for user id
String randomString = RandomStringUtils.randomAlphabetic(7)
String newValue = "User_name_" + randomString

// Define the XPath
String xpath_for_user_name = "//label[text()='Username']/ancestor::div[@class='oxd-input-group__label-wrapper']/following-sibling::div//input[@class='oxd-input oxd-input--active']"

// Locate the input field using the XPath
TestObject inputField = new TestObject().addProperty("xpath", ConditionType.EQUALS, xpath_for_user_name)

// Use JavaScript to set the value and trigger the input event
WebUI.executeJavaScript(
    """
    var element = document.evaluate("${xpath_for_user_name}", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (element) {
        element.value = '${newValue}';
        var event = new Event('input', { bubbles: true });
        element.dispatchEvent(event);
    }
    """, 
    null
)

WebUI.delay(1)


GlobalVariable.User_name = WebUI.getAttribute(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_username'),'value')

System.out.println(GlobalVariable.User_name)

System.out.println(GlobalVariable.Password)


//WebUI.setText(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_username'), 'User_name_' + RandomStringUtils.randomAlphabetic(6))

GlobalVariable.First_name = WebUI.getAttribute(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_Employee_first_name'),'value')

System.out.println(GlobalVariable.First_name)


GlobalVariable.Middle_name = WebUI.getAttribute(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_middle_name'),'value')

System.out.println(GlobalVariable.Middle_name)


GlobalVariable.Last_name = WebUI.getAttribute(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_last_name'),'value')

System.out.println(GlobalVariable.Last_name)


GlobalVariable.employee_id = WebUI.getAttribute(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_id'),'value')

System.out.println(GlobalVariable.employee_id)

WebUI.delay(3)

WebUI.click(findTestObject('Object Repository/ADDING_A_NEW_EMPLOYEE/Add_employee_save_button'))

WebUI.delay(2)
