package com.cucumber.framework.utility;


import java.util.Properties;

import java.io.IOException;
import java.util.Properties;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.io.File;

import java.io.FileInputStream;

import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SendEmailUtility {

	public static void main(String[] args) throws Exception {

		readExcel("D:\\", "TestData.xlsx", "Sheet1");

	}

	public static void sendMail(String from_address, String pwd, String to_address, String subject,String email_body) {
		final String username = from_address;
		final String password = pwd;

		Properties props = new Properties();
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.host", "outlook.office365.com");
		props.put("mail.smtp.port", "587");

		Session session = Session.getInstance(props, new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(username, password);
			}
		});

		try {

			Message message = new MimeMessage(session);
			message.setFrom(new InternetAddress(username));
			message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(to_address));
			message.setSubject(subject);
			message.setText(email_body);

			Transport.send(message);

			System.out.println("Mail sent");

		} catch (MessagingException e) {
			throw new RuntimeException(e);
		}

	}

	public static void readExcel(String filePath, String fileName, String sheetName) throws IOException {

		// Create an object of File class to open xlsx file

		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file

		FileInputStream inputStream = new FileInputStream(file);

		Workbook Workbook = null;

		// Find the file extension by splitting file name in substring and getting only
		// extension name

		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file

		if (fileExtensionName.equals(".xlsx")) {

			// If it is xlsx file then create object of XSSFWorkbook class

			Workbook = new XSSFWorkbook(inputStream);

		}

		// Check condition if the file is xls file

		else if (fileExtensionName.equals(".xls")) {

			// If it is xls file then create object of HSSFWorkbook class

			Workbook = new HSSFWorkbook(inputStream);

		}

		// Read sheet inside the workbook by its name

		Sheet Sheet = Workbook.getSheet(sheetName);

		// Find number of rows in excel file

		int rowCount = Sheet.getLastRowNum() - Sheet.getFirstRowNum();

		// Create a loop over all the rows of excel file to read it

		for (int i = 1; i <= rowCount; i++) {

			Row row = Sheet.getRow(i);

			// Create a loop to print cell values in a row

			String from_address=row.getCell(0).getStringCellValue();
			String from_address_pwd=row.getCell(1).getStringCellValue();
			String to_address=row.getCell(2).getStringCellValue();
			String subject = row.getCell(3).getStringCellValue();
			String email_body = row.getCell(4).getStringCellValue();
			sendMail(from_address,from_address_pwd,to_address,subject,email_body); 
			System.out.println(i + " Mail Sent");
		}
		Workbook.close();
		inputStream.close();
	}

}
