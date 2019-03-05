package com.demo.JDBCExcel;

public class Employee {
	private String Name;
	private String Email;
	private String Salary;
	
	public Employee(String name, String email, String esalary) {
		super();
		Name = name;
		Email = email;
		Salary = esalary;
	}
	public String getName() {
		return Name;
	}
	public void setName(String name) {
		Name = name;
	}
	public String getEmail() {
		return Email;
	}
	public void setEmail(String email) {
		Email = email;
	}
	public String getSalary() {
		return Salary;
	}
	public void setSalary(String salary) {
		Salary = salary;
	}
	
	
}
