package com.app.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class TransformerImpl implements Transformer {

    public void saveToExcel(List<Person> personList) {
        Workbook excelDocument = new XSSFWorkbook();
        Sheet sheet = excelDocument.createSheet("Perons Sheet");
        Row row=sheet.createRow(0);
        Cell cellName=row.createCell(0);
        cellName.setCellValue("NAME");
        Cell cellAge = row.createCell(1);
        cellAge.setCellValue("AGE");
        int rowIndex=1;
        for(Person p: personList){
            Row personRow=sheet.createRow(rowIndex++);
            Cell personNameCell=personRow.createCell(0);
            personNameCell.setCellValue(p.getName());
            Cell personAgeCell=personRow.createCell(1);
            personAgeCell.setCellValue(p.getAge());
        }


        File file = new File("persons.xls");
        try(OutputStream out = new FileOutputStream(file)){
            excelDocument.write(out);
        } catch(Exception ex){
             ex.printStackTrace();
        }



    }

    public List<Person> fromExcel(String path) {
        try {
            List<Person> list = new ArrayList();
            Workbook excelDocument = new XSSFWorkbook(path);
            Sheet sheet=excelDocument.getSheetAt(0);
            boolean isHeader = true;
            for(Row row: sheet){
                if(isHeader){
                    isHeader=false;
                    continue;
                }
                Cell nameCell=row.getCell(0);
                Cell ageCell = row.getCell(1);
                Person person = new Person();
                person.setName(nameCell.getStringCellValue());
                person.setAge((int)ageCell.getNumericCellValue());
                list.add(person);
           }
           return list;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
