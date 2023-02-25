// updated on 25.02.23 @ 09.00 hrs
package weeklysoa;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.InetAddress;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
public class SoaMail {
	public static void sendEmailWithAttachments(String host, String port, final String userName, final String password,
			String toAddress, String subject, String message, String[] attachFiles)
			throws AddressException, MessagingException {
		// sets SMTP server properties
		Properties properties = new Properties();
		properties.put("mail.smtp.host", host);
		properties.put("mail.smtp.port", port);
		properties.put("mail.smtp.auth", "true");
		properties.put("mail.smtp.starttls.enable", "true");
	 	properties.setProperty("mail.smtp.ssl.protocols", "TLSv1.2");
	 	properties.put("mail.smtp.isSSL","true");
		properties.put("mail.debug", "true");
		properties.put("mail.user", userName);
		properties.put("mail.password", password);
		// creates a new session with an authenticator
		Authenticator auth = new Authenticator() {
			public PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(userName, password);
			}
		};
		Session session = Session.getInstance(properties, auth);
		// creates a new e-mail message
		Message msg = new MimeMessage(session);
		msg.setFrom(new InternetAddress(userName));
		InternetAddress[] toAddresses = { new InternetAddress(toAddress) };
		msg.setRecipients(Message.RecipientType.TO, toAddresses);
		msg.setSubject(subject);
		msg.setSentDate(new Date());
		// creates message part
		MimeBodyPart messageBodyPart = new MimeBodyPart();
		messageBodyPart.setContent(message, "text/html");
		// creates multi-part
		Multipart multipart = new MimeMultipart();
		multipart.addBodyPart(messageBodyPart);
		// adds attachments
		if (attachFiles != null && attachFiles.length > 0) {
			for (String filePath : attachFiles) {
				MimeBodyPart attachPart = new MimeBodyPart();
				try {
					attachPart.attachFile(filePath);
				} catch (IOException ex) {
					ex.printStackTrace();
				}
				multipart.addBodyPart(attachPart);
			}
		}
		// sets the multi-part as e-mail's content
		msg.setContent(multipart);
		// sends the e-mail
		Transport.send(msg);
	}
	public static void main(String[] args) throws IOException, Exception, Exception {
		Map<Integer, String> add1 = new HashMap<Integer, String>();
		Map<Integer, String> add2 = new HashMap<Integer, String>();
		FileInputStream xlsfile = new FileInputStream("d:/debtors/drs 2022-23/Master.xlsx");
		Workbook xls = WorkbookFactory.create(xlsfile);
		Sheet sht = xls.getSheetAt(2);
		for (Row r : sht) {
			if (r.getRowNum() > 0) {
				add1.put((int) r.getCell(0).getNumericCellValue(), r.getCell(2).getStringCellValue());
				add2.put((int) r.getCell(0).getNumericCellValue(), r.getCell(3).getStringCellValue());
			}
		}
		xlsfile.close();
		int code = 0;
		File[] soa = new File("d://mails/").listFiles();
		for (File s : soa) {
			PDDocument soamail = Loader.loadPDF(s);
			{
				PDFTextStripper soatext = new PDFTextStripper();
				soatext.setStartPage(1);
				soatext.setEndPage(1);
				String pdfFileInText = soatext.getText(soamail);
				String lines[] = pdfFileInText.split("\\r?\\n");
				code = Integer.parseInt(lines[4].substring(lines[4].indexOf(":") + 2));
				System.out.println(lines[4].substring(lines[4].indexOf(":") + 2) + " ~ "
						+ lines[5].substring(lines[5].indexOf(":") + 2) + " ~ "
						+ lines[16].substring(lines[16].indexOf(":") + 2, 43) + " ~ " + add1.get(code) + " ~ "
						+ add2.get(code));
			}
		}
	
		//System.out.println( InetAddress.getLocalHost().getHostName());
		
		boolean t = false;
		String host = "smtp.gmail.com";
		String port = "587";
		String mailFrom = "ndp15143@gmail.com";
		String password = "conornnanbqzpqgu";
		if (t) {
			host = "localhost";
			port = "135";
			mailFrom = "dhandapani@futureconsumer.in";
			password = "Future@2023";
		}
		
		System.out.println(InetAddress.getLocalHost().getHostName());
		 
		// message info
		String mailTo = "ndp15143@yahoo.com";
		mailTo="dhandapani@futureconsumer.in";
		String subject = "New email with attachments no 333  " + mailFrom;
		String message = "I have some attachments for you.";
		// attachments
		String[] attachFiles = new String[1];
		attachFiles[0] = "d:/tnstore/ledger.xlsx";
		//attachFiles[1] = "d:/tnstore/ledger.xlsx";
		//attachFiles[2] = "d:/tnstore/ledger.xlsx";
		try {
			sendEmailWithAttachments(host, port, mailFrom, password, mailTo, subject, message, attachFiles);
			System.out.println("Email sent. ..... dhandapani");
		} catch (Exception ex) {
			System.out.println("Could not send email.");
			ex.printStackTrace();
		}
	}
}
