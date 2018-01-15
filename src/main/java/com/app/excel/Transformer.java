package com.app.excel;

import java.util.List;

public interface Transformer {
    public void saveToExcel(List<Person> personList);
    public List<Person> fromExcel(String path);
}
