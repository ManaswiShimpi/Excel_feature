package com.lti.service;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.lti.model.Employee;
import com.lti.repository.EmployeeRepository;

@Service
public class EmployeeService {
	@Autowired
	private EmployeeRepository employeeRepository;

//create
	public Employee create(String empid, String name, String department, float salary) {
		return employeeRepository.save(new Employee(empid, name, department, salary));

	}

//retrive
	
	public List<Employee> getAllBySalary() throws IOException {

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Sample sheet");
		Map<String, Object[]> data = new HashMap<String, Object[]>();

		List<Employee> employee = employeeRepository.findAll().stream().filter(emp -> emp.getSalary() > 20000)
				.collect(Collectors.toList());

		Row headerRow = sheet.createRow(0);

		Cell headerCell = headerRow.createCell(1);
		headerCell.setCellValue("Emp Id");

		headerCell = headerRow.createCell(2);
		headerCell.setCellValue("Name");

		headerCell = headerRow.createCell(3);
		headerCell.setCellValue("Department");

		headerCell = headerRow.createCell(4);
		headerCell.setCellValue("salary");
		int rowCount = 1;
		// Start
		for (Employee emp : employee) {
			// while (employee.equals(employeeRepository)) {
			Row row = sheet.createRow(rowCount++);
			// Row row = sheet.createRow(1);
			Cell cell = row.createCell(1);
			cell.setCellValue(emp/* .get(0) */.getEmpid());
			cell = row.createCell(2);
			cell.setCellValue(emp/* .get(0) */.getName());
			cell = row.createCell(3);
			cell.setCellValue(emp/* .get(0) */.getDepartment());
			cell = row.createCell(4);
			cell.setCellValue(emp/* .get(0) */.getSalary());
			// End
		}
		FileOutputStream out = null;

		try {

		                   out = new FileOutputStream(new   File("C:\\Users\\91976\\Desktop\\new.xls"));

		                     workbook.write(out);

		                     //don’t close the stream here.

		              } catch (FileNotFoundException e) {

		                     e.printStackTrace();

		              } catch (IOException e) {

		                     e.printStackTrace();

		              }  
		finally{

		                     out.close();

		}

		 

		return (List<Employee>) employeeRepository.findAll().stream().filter(emp -> emp.getSalary() > 20000)
				.collect(Collectors.toList());
	}
	 
	public List<Employee> getAll() {
		return employeeRepository.findAll();
	}

	public Employee getByEmpid(String empid) {
		return employeeRepository.findByEmpid(empid);
	}

//update
	public Employee update(String empid, String name, String department, float salary) {
		Employee p = employeeRepository.findByEmpid(empid);
		p.setName(name);
		p.setDepartment(department);
		p.setSalary(salary);
		return employeeRepository.save(p);
	}

//Delete
	public void deleteAll() {
		employeeRepository.deleteAll();
	}

	public void delete(String empid) {
		Employee p = employeeRepository.findByEmpid(empid);
		employeeRepository.delete(p);
	}
}
