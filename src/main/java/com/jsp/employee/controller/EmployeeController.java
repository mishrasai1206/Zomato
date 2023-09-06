package com.jsp.employee.controller;

import java.io.IOException;
import java.util.List;

import javax.mail.MessagingException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.jsp.employee.entity.Employee;
import com.jsp.employee.service.EmployeeService;
import com.jsp.employee.utility.MessageData;
import com.jsp.employee.utility.ResponseStructure;

@RestController
@RequestMapping("/employee")
public class EmployeeController {
	@Autowired
	private EmployeeService employeeService;

	@PostMapping("/printFile")
	public ResponseEntity<String> printExtractFromExcel(MultipartFile multipartFile) throws IOException {
		return employeeService.printExtractFromExcel(multipartFile);
	}

	@PostMapping("/extractFromExcel")
	public ResponseEntity<ResponseStructure<List<Employee>>> extractFromExcel(MultipartFile multipartFile)
			throws IOException {
		return employeeService.extractFromExcel(multipartFile);
	}

	@PostMapping("/uploadIntoExcel")
	public ResponseEntity<String> uploadToExcel(String filePath) throws IOException {
		return employeeService.uploadToExcel(filePath);
	}

	@PostMapping("/uploadFromCsv")
	public ResponseEntity<String> extractFromCsv(MultipartFile multipartFile) throws IOException {
		return employeeService.extractFromCsv(multipartFile);
	}

	@PostMapping("/uploadToCsv")
	public ResponseEntity<String> uploadToCsv(String filePath) throws IOException {
		return employeeService.uploadToCsv(filePath);
	}

	@PostMapping("/sendMail")
	public ResponseEntity<String> sendMailToStudents(@RequestBody MessageData messageData) {
		return employeeService.sendMailToStudents(messageData);

	}

	@CrossOrigin
	@PostMapping("/sendMimeMail")
	public ResponseEntity<String> sendMimeMailToStudents(@RequestBody MessageData messageData)
			throws MessagingException {
		return employeeService.sendMimeMailToStudents(messageData);

	}
}
