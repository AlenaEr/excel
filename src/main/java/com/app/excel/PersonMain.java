package com.app.excel;

import java.util.ArrayList;
import java.util.List;

public class PersonMain {
    public static void main(String[] args) {
        List<Person> personList=new ArrayList<Person>();
       for (int i=0;i<3;i++){
           Person p = new Person();
           p.setAge(i);
           p.setName("Name"+i);
           personList.add(p);
       }
       Transformer transformer = new TransformerImpl();
        transformer.saveToExcel(personList);
        List<Person> personsResultList=transformer.fromExcel("persons.xls");
        for(Person p: personsResultList){
            System.out.println(p);
        }
    }
}
