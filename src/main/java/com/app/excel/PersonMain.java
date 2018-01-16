package com.app.excel;

import java.util.ArrayList;
import java.util.List;

public class PersonMain {
    public static void main(String[] args) {
//        List<Person> personList = new ArrayList<Person>();
//        for (int i = 0; i < 3; i++) {
//            Person p = new Person();
//            p.setAge(i);
//            p.setName("Name" + i);
//            personList.add(p);
//        }
//        Transformer transformer = new TransformerImpl();
//        transformer.saveToExcel(personList);
//        List<Person> personsResultList = transformer.fromExcel("persons.xls");
//        for (Person p : personsResultList) {
//            System.out.println(p);
//        }

        System.out.println("-----------------------Emploee------------------------");

        List<Employee> empl = new ArrayList<Employee>();
        for (int i = 0; i < 3; i++) {
            Employee em = new Employee();
            em.setPhone(i);
            em.setName("Name" + i);
            em.setDepartment("department");
            empl.add(em);
        }
        EmployeeTransform employeeTransform = new EmployeeImpl();
        employeeTransform.saveToExcel(empl);
        List<Employee> empRes = employeeTransform.fromExcel("employee.xls");
        for (Employee p : empRes) {
            System.out.println(p);
        }
    }
}
