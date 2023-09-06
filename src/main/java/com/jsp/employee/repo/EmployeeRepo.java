package com.jsp.employee.repo;

import org.springframework.data.jpa.repository.JpaRepository;

import com.jsp.employee.entity.Employee;

public interface EmployeeRepo extends JpaRepository<Employee, Integer>{

}
