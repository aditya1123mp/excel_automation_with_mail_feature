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
import javax.activation.DataHandler
import javax.activation.FileDataSource

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
    println('Searching for emails on or after: ' + specificDate)

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

            println('Email sender: ' + email)
            println('Full From field: ' + sender)
            println('Subject: ' + message.getSubject())
            println('Sent Date: ' + message.getSentDate())

            // Check if it matches the target address
            if (email.equalsIgnoreCase('aditya_T2_WORK@outlook.com')) {
                println('Match found for: ' + email)

                Object content = message.getContent()
                println('Content Type: ' + message.getContentType())

                if (content instanceof Multipart) {
                    Multipart multipart = (Multipart) content

                    for (int i = 0; i < multipart.getCount(); i++) {
                        BodyPart bodyPart = multipart.getBodyPart(i)
                        println('Part Content Type: ' + bodyPart.getContentType())

                        // Handle attachments with generic MIME type and .xlsx or .xls file extension
                        if (bodyPart.getContentType().toLowerCase().contains("application/octet-stream") &&
                            (bodyPart.getFileName().endsWith(".xlsx") || bodyPart.getFileName().endsWith(".xls"))) {

                            // Specify where to save the attachment
                            String saveDirectory = 'C:\\Users\\DELL\\OneDrive\\Desktop\\excel_folder'
                            String fileName = bodyPart.getFileName()

                            if (fileName == null || fileName.trim().isEmpty()) {
                                fileName = 'default_excel_file.xlsx'
                            }

                            println('Attachment file name: ' + fileName)

                            File fileToSave = new File(saveDirectory, fileName)
                            println('Saving file to: ' + fileToSave.getAbsolutePath())

                            try {
                                // Save the attachment
                                bodyPart.saveFile(fileToSave)
                                println('Excel file saved successfully: ' + fileToSave.getAbsolutePath())
                            } catch (IOException e) {
                                println('Failed to save file: ' + e.getMessage())
                            }
                        } else {
                            println('No Excel attachment found in this part. MIME type: ' + bodyPart.getContentType())
                        }
                    }
                } else {
                    println('Email content is not multipart, unable to process attachment.')
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
} catch (IOException e) {
    e.printStackTrace()
}
