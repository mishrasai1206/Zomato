package com.jsp.employee.serviceimpl;

import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.jsp.employee.entity.Employee;
import com.jsp.employee.repo.EmployeeRepo;
import com.jsp.employee.service.EmployeeService;
import com.jsp.employee.utility.MessageData;
import com.jsp.employee.utility.ResponseStructure;

@Service
public class EmployeeServiceImpl implements EmployeeService {

	@Autowired
	private EmployeeRepo employeeRepo;

	@Autowired
	private JavaMailSender javaMailSender;

	@Override
	public ResponseEntity<ResponseStructure<Employee>> saveEmployee(Employee employee) {
		return null;
	}

	@Override
	public ResponseEntity<String> printExtractFromExcel(MultipartFile multipartFile) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook(multipartFile.getInputStream());
		for (Sheet sheet : workbook) {
			for (Row row : sheet) {
				System.out.println(row.getCell(0).getStringCellValue());
				System.out.println(row.getCell(1).getStringCellValue());
				System.out.println(row.getCell(2).getStringCellValue());
			}
		}
		return new ResponseEntity<String>("Data Displayed on the console!!", HttpStatus.FOUND);
	}

	@Override
	public ResponseEntity<ResponseStructure<List<Employee>>> extractFromExcel(MultipartFile multipartFile)
			throws IOException {
		List<Employee> employees = new ArrayList<>();
		XSSFWorkbook workbook = new XSSFWorkbook(multipartFile.getInputStream());
		for (Sheet sheet : workbook) {
			for (Row row : sheet) {
				Employee employee = new Employee();
				employee.setEmployeeName(row.getCell(0).getStringCellValue());
				employee.setEmployeeEmail(row.getCell(1).getStringCellValue());
				employee.setEmployeePassword(row.getCell(2).getStringCellValue());
				employeeRepo.save(employee);
				employees.add(employee);

			}
		}
		ResponseStructure<List<Employee>> responseStructure = new ResponseStructure<>();
		responseStructure.setHttpStatusCode(HttpStatus.CREATED.value());
		responseStructure.setMessage("Students saved Successfully!!");
		responseStructure.setA(employees);

		return new ResponseEntity<ResponseStructure<List<Employee>>>(responseStructure, HttpStatus.CREATED);

	}

	@Override
	public ResponseEntity<String> uploadToExcel(String filePath) throws IOException {
		List<Employee> employees = employeeRepo.findAll();

		if (employees.isEmpty()) {
			return null;
		} else {
			// create a new WorkBook
			XSSFWorkbook workbook = new XSSFWorkbook();

			// create a new sheet
			Sheet sheet = workbook.createSheet("Employee Data");

			// create headers
			Row headerRow = sheet.createRow(0);
			headerRow.createCell(0).setCellValue("ID");
			headerRow.createCell(1).setCellValue("Name");
			headerRow.createCell(2).setCellValue("Email");
			headerRow.createCell(3).setCellValue("Password");

			// populate rows with employee data
			int rowNum = 1;
			for (Employee employee : employees) {
				Row row = sheet.createRow(rowNum++);
				row.createCell(0).setCellValue(employee.getEmployeeId());
				row.createCell(1).setCellValue(employee.getEmployeeName());
				row.createCell(2).setCellValue(employee.getEmployeeEmail());
				row.createCell(3).setCellValue(employee.getEmployeePassword());
			}

			// save the workbook to a file
			FileOutputStream fileOutputStream = new FileOutputStream(filePath);
			workbook.write(fileOutputStream);

			// close the workbook
			workbook.close();
			return new ResponseEntity<String>("Data uploaded Successfully!!", HttpStatus.OK);
		}
	}

	@Override
	public ResponseEntity<String> extractFromCsv(MultipartFile multipartFile) throws IOException {
		// Create a CSVReader object.
		Reader reader = new InputStreamReader(multipartFile.getInputStream());
		CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT);
		List<CSVRecord> records = csvParser.getRecords();

		for (CSVRecord record : records) {
			Employee employee = new Employee();
			employee.setEmployeeName(record.get(0));
			employee.setEmployeeEmail(record.get(1));
			employee.setEmployeePassword(record.get(2));
			employeeRepo.save(employee);
		}
		csvParser.close();
		return new ResponseEntity<String>("Data saved from csv!!", HttpStatus.OK);
	}

	@Override
	public ResponseEntity<String> uploadToCsv(String filePath) throws IOException {
		List<Employee> employees = employeeRepo.findAll();

		if (employees.isEmpty()) {
			return null;
		} else {
			FileWriter fileWriter = new FileWriter(filePath);
			CSVPrinter csvPrinter = new CSVPrinter(fileWriter, CSVFormat.DEFAULT);

			for (Employee employee : employees) {
				csvPrinter.printRecord(employee.getEmployeeId(), employee.getEmployeeName(),
						employee.getEmployeeEmail(), employee.getEmployeePassword());
			}
			csvPrinter.flush();
			csvPrinter.close();
		}
		return new ResponseEntity<String>("Data inserted inside csv!!", HttpStatus.OK);
	}

	@Override
	public ResponseEntity<String> sendMailToStudents(MessageData messageData) {
		SimpleMailMessage message = new SimpleMailMessage();
		message.setTo(messageData.getTo());
		message.setSubject(messageData.getSubject());
		message.setText(
				messageData.getText() + "\n\n" + messageData.getSenderName() + ",\n" + messageData.getSenderAddress());
		message.setSentDate(new Date());
		javaMailSender.send(message);
		return new ResponseEntity<String>("Mail sent Successfully!!", HttpStatus.OK);
	}

	@Override
	public ResponseEntity<String> sendMimeMailToStudents(MessageData messageData) throws MessagingException {
		MimeMessage mimeMessage = javaMailSender.createMimeMessage();
		MimeMessageHelper helper = new MimeMessageHelper(mimeMessage, true);
//		helper.setFrom(new InternetAddress("mishrasaicharan@gmail.com"));
		helper.setTo(messageData.getTo());
		helper.setSubject(messageData.getSubject());
		helper.setSentDate(new Date());
		String emailBody = messageData.getText() + "<br><br><h4>Thanks & Regards</h4>" + "<h4>"
				+ messageData.getSenderName() + ",<br>" + messageData.getSenderAddress() + "</h4>";

		helper.setText(emailBody, true);
		javaMailSender.send(mimeMessage);
		return new ResponseEntity<String>("Mail Sent Successfully!!", HttpStatus.OK);
	}

}
