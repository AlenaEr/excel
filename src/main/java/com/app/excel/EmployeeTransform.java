package com.app.excel;

import java.util.List;

public interface EmployeeTransform {
    public void saveToExcel(List<Employee> employees);

    public List<Employee> fromExcel(String path);
}
