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
import java.util.Properties as Properties
import javax.mail.search.FlagTerm as FlagTerm
import javax.mail.search.SearchTerm as SearchTerm
import javax.mail.search.SentDateTerm as SentDateTerm
import javax.mail.search.AndTerm as AndTerm
import java.util.Calendar as Calendar
import java.util.Date as Date

WebUI.delay(3)

// Define email properties for reading from Gmail
Properties props = new Properties()

props.put('mail.store.protocol', 'imaps')
props.put('mail.imaps.host', 'imap.gmail.com')
props.put('mail.imaps.port', '993')
props.put('mail.imaps.ssl.enable', 'true')

// Define email credentials
String username = GlobalVariable.receiver_username
String password = GlobalVariable.receiver_password // App password

// Create session
Session session = Session.getInstance(props, null)

try {
    // Connect to the email store
    Store store = session.getStore('imaps')
    store.connect(username, password)

    // Open the inbox folder
    Folder inbox = store.getFolder('INBOX')
    inbox.open(Folder.READ_ONLY)

    // Set the Calendar to the specific date 09/08/2024 without time
    Calendar cal = Calendar.getInstance()
    cal.set(Calendar.YEAR, 2024)
    cal.set(Calendar.MONTH, Calendar.AUGUST)
    cal.set(Calendar.DAY_OF_MONTH, 9)
    cal.set(Calendar.HOUR_OF_DAY, 0)
    cal.set(Calendar.MINUTE, 0)
    cal.set(Calendar.SECOND, 0)
    cal.set(Calendar.MILLISECOND, 0)

    // Convert to Date
    Date specificDate = cal.getTime()

    // Create a date search term for emails sent on or after 09/08/2024
    SearchTerm newerThan = new SentDateTerm(SentDateTerm.GE, specificDate)

    // Combine the unseen filter with the date filter
    SearchTerm unseenAndNewer = new AndTerm(new FlagTerm(new Flags(Flags.Flag.SEEN), false), newerThan)

    // Search for unseen and recent emails
    Message[] messages = inbox.search(unseenAndNewer)
    println('Total messages found: ' + messages.length)

    if (messages.length > 0) {
        for (Message message : messages) {
            Address[] froms = message.getFrom()
            String sender = (froms[0]).toString()
            String email = sender.replaceAll('.*<|>', '').trim()

            // Print all relevant information for debugging
            println('Email sender: ' + email)
            println('Full From field: ' + sender)
            println('Subject: ' + message.getSubject())
            println('Sent Date: ' + message.getSentDate())

            // Check if it matches the target address
            if (email.equalsIgnoreCase('aditya_T2_WORK@outlook.com')) {
                println('Match found for: ' + email)

                // Print content type for debugging
                println('Content Type: ' + message.getContentType())

                Object content = message.getContent()
                String textContent = ''

                if (content instanceof String) {
                    textContent = content.toString() // Debugging output
                } else if (content instanceof Multipart) {
                    Multipart multipart = (Multipart) content // Corrected type casting

                    for (int i = 0; i < multipart.getCount(); i++) {
                        BodyPart bodyPart = multipart.getBodyPart(i)

                        println('Part Content Type: ' + bodyPart.getContentType())

                        if (bodyPart.isMimeType('text/plain')) {
                            textContent = bodyPart.getContent().toString()
                            break
                        }
                    }
                }
                
                println('Extracted Content: ' + textContent)

                // Extract username and password
                def usernameMatcher = textContent =~ 'Username:\\s*(\\S+)'
                def passwordMatcher = textContent =~ 'Password:\\s*(\\S+)'

                if (usernameMatcher.find() && passwordMatcher.find()) {
                    String extractedUsername = usernameMatcher.group(1)
                    String extractedPassword = passwordMatcher.group(1)

                    println('Extracted Username: ' + extractedUsername)
                    println('Extracted Password: ' + extractedPassword)

                    // You can store these in global variables if needed
                    GlobalVariable.extractedUsername = extractedUsername
                    GlobalVariable.extractedPassword = extractedPassword
                } else {
                    println('Username or Password not found in the email content.')
                }
                
                break
            } else {
                println('No match for: ' + email)
            }
        }
    } else {
        println('No new emails found.')
    }
    
    inbox.close(false)
    store.close()
}
catch (MessagingException e) {
    e.printStackTrace()
} 
