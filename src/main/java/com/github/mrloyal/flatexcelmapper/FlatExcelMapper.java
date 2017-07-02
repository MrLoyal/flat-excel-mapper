package com.github.mrloyal.flatexcelmapper;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.util.List;

public class FlatExcelMapper {

    public <T extends Object>List<T> read(String fileName, int sheetIndex, Class<T> clazz){
        List<T> list = null;

        return list;
    }

    public <T extends Object>List<T> read(File file, int sheetIndex, Class<T> clazz){
        List<T> list = null;

        return list;
    }

    public <T extends Object>List<T> read(XSSFSheet sheet, Class<T> clazz){
        List<T> list = null;

        return list;
    }

    public <T extends Object> T readRow(XSSFSheet sheet, int humanReadableRow, Class<T> clazz) throws IllegalAccessException, InstantiationException {
        T object = clazz.newInstance();

        return object;
    }

    public <T extends Object> T readRow(XSSFRow excelRow, Class<T> clazz){
        return null;
    }
}
