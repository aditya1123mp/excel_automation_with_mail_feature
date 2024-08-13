import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

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
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows

import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys

import javax.mail.*
import javax.mail.internet.*
import java.util.Properties

// Define email properties
Properties props = new Properties()
props.put("mail.smtp.host", "smtp-mail.outlook.com") // Replace with your SMTP server
props.put("mail.smtp.port", "587") // Replace with your SMTP port
props.put("mail.smtp.auth", "true")
props.put("mail.smtp.starttls.enable", "true")

// Email credentials (assuming they are predefined)
final String emailUsername = GlobalVariable.user_name_for_sender // Replace with your email
final String emailPassword = GlobalVariable.sender_password // Replace with your email password

// Create session
Session session = Session.getInstance(props, new Authenticator() {
    protected PasswordAuthentication getPasswordAuthentication() {
        return new PasswordAuthentication(emailUsername, emailPassword)
    }
})

try {
    // Create a default MimeMessage object
    Message message = new MimeMessage(session)

    // Set From: header field
    message.setFrom(new InternetAddress(emailUsername))

    // Set To: header field
    message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(GlobalVariable.receiver_username)) // Replace with recipient email

    // Set Subject: header field
    message.setSubject("Automation Test Results")

    // Set the content of the email to include global variables
    String emailBody = //"Test Result: " + GlobalVariable.textboxvalue + 
                       "\nUsername: " + GlobalVariable.User_name + 
                       "\nPassword: " + GlobalVariable.Password
    message.setText(emailBody)

    // Send message
    Transport.send(message)

    println("Email sent successfully!")

} catch (MessagingException e) {
    e.printStackTrace()
}