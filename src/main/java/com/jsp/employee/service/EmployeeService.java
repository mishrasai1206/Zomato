package com.jsp.employee.service;

import java.io.IOException;
import java.util.List;

import javax.mail.MessagingException;

import org.springframework.http.ResponseEntity;
import org.springframework.web.multipart.MultipartFile;

import com.jsp.employee.entity.Employee;
import com.jsp.employee.utility.MessageData;
import com.jsp.employee.utility.ResponseStructure;

public interface EmployeeService {
	public ResponseEntity<ResponseStructure<Employee>> saveEmployee(Employee employee);

	public ResponseEntity<String> printExtractFromExcel(MultipartFile multipartFile) throws IOException;

	public ResponseEntity<ResponseStructure<List<Employee>>> extractFromExcel(MultipartFile multipartFile)
			throws IOException;

	public ResponseEntity<String> uploadToExcel(String filePath) throws IOException;

	public ResponseEntity<String> extractFromCsv(MultipartFile multipartFile) throws IOException;

	public ResponseEntity<String> uploadToCsv(String filePath) throws IOException;
	
	public ResponseEntity<String> sendMailToStudents(MessageData messageData);
	
	public ResponseEntity<String> sendMimeMailToStudents(MessageData messageData) throws MessagingException;
}




