package com.app.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class EmployeeImpl implements EmployeeTransform {
    @Override
    public void saveToExcel(List<Employee> employees) {
        Workbook excelDocument = new XSSFWorkbook();
        Sheet sheet = excelDocument.createSheet("Employee Sheet");
        Row row = sheet.createRow(0);
        Cell cellName = row.createCell(0);
        cellName.setCellValue("name");
        Cell cellDepartment = row.createCell(1);
        cellDepartment.setCellValue("department");
        Cell cellPhone = row.createCell(2);
        cellPhone.setCellValue("phone");
        int rowNum = 1;
        for (Employee em : employees) {
            Row rowEmpl = sheet.createRow(rowNum++);
            Cell empNameCell = rowEmpl.createCell(0);
            empNameCell.setCellValue(em.getName());
            Cell depCell = rowEmpl.createCell(1);
            depCell.setCellValue(em.getDepartment());
            Cell phoneNum = rowEmpl.createCell(2);
            phoneNum.setCellValue(em.getPhone());
        }
        File file = new File("employee.xls");
        try {
            OutputStream outputStream = new FileOutputStream(file);
            excelDocument.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public List<Employee> fromExcel(String path) {
        try {
            List<Employee> employees = new ArrayList();
            InputStream in = new FileInputStream(new File("D:\\excel\\employee.xls"));
            Workbook excelBook = new XSSFWorkbook(in);

            Sheet sheet = excelBook.getSheet("Employee Sheet");

            boolean header = true;
            for (Row row : sheet) {
                if (header) {
                    header = false;
                    continue;
                }
                Cell empName = row.getCell(0);
                Cell depart = row.getCell(1);
                Cell phone = row.getCell(2);
                Employee employee = new Employee();
                employee.setName(empName.getStringCellValue());
                employee.setDepartment(depart.getStringCellValue());
                employee.setPhone((int) phone.getNumericCellValue());
                employees.add(employee);
            }
            return employees;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
