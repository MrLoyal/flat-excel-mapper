package com.github.mrloyal.flatexcelmapper;

import com.github.mrloyal.flatexcelmapper.annotation.DateSourceType;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelColumn;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelDate;
import com.github.mrloyal.flatexcelmapper.exception.AnnotationAttributeMissingException;
import com.github.mrloyal.flatexcelmapper.exception.DataTypeNotSupportedException;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

public class FlatExcelMapper {

    private DataFormatter dataFormatter = new DataFormatter();

    public <T extends Object>List<T> read(String fileName, int sheetIndex, Class<T> clazz) throws Exception {
        File file = new File(fileName);
        return read(file, sheetIndex, clazz);
    }

    public <T extends Object>List<T> read(File file, int sheetIndex, Class<T> clazz)
            throws IOException, IntrospectionException, InstantiationException,
            DataTypeNotSupportedException, IllegalAccessException, InvocationTargetException, ParseException {
        List<T> list = null;
        XSSFWorkbook workbook;
        FileInputStream fis = null;
        IOException ioException = null;
        XSSFSheet sheet;

        try {
            fis = new FileInputStream(file);
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(sheetIndex);

            Iterator<Row> iterator = sheet.rowIterator();
            while (iterator.hasNext()){
                XSSFRow row = (XSSFRow) iterator.next();
                XSSFCell cell = row.getCell(1);
                // System.out.println("cell: " + cell);
                readRow(row, clazz);
            }

            readRow(sheet, 5, clazz);
        } catch (IOException e){
            ioException = e;
        } finally {
            if (fis != null){
                fis.close();
            }
        }

        if (ioException != null){
            throw ioException;
        }
        return list;
    }

    public <T extends Object>List<T> read(XSSFWorkbook workbook, int sheetIndex, Class<T> clazz){
        return null;
    }


    public <T extends Object>List<T> read(XSSFSheet sheet, Class<T> clazz){
        List<T> list = null;

        return list;
    }

    public <T extends Object> T readRow(XSSFSheet sheet, int humanReadableRow, Class<T> clazz)
            throws IllegalAccessException, InstantiationException, IntrospectionException,
            DataTypeNotSupportedException, InvocationTargetException, ParseException {
        T object = clazz.newInstance();
        XSSFRow row = sheet.getRow(humanReadableRow-1);
        XSSFCell cell;
        for (Method method : clazz.getDeclaredMethods()){
            if (method.isAnnotationPresent(ExcelColumn.class)){
                ExcelColumn colAnn = method.getAnnotation(ExcelColumn.class);
                boolean nullable = colAnn.nullable();
                String colName = colAnn.name();
                int colIndex = CellReference.convertColStringToIndex(colName);
                cell = row.getCell(colIndex);
                PropertyDescriptor pd = getPropertyDescriptor(method);

                if (method.isAnnotationPresent(ExcelDate.class)){
                    ExcelDate dateAnn = method.getAnnotation(ExcelDate.class);
                    DateSourceType type = dateAnn.type();
                    String pattern = null;
                    if (type == DateSourceType.STRING){
                        pattern = dateAnn.format();
                    }
                    setCellValueToDateProperty(object, pd, cell, nullable, type, pattern);
                } else {
                    setCellValueToProperty(object, pd, cell, nullable);
                }
            }
        }
        return object;
    }

    private void setCellValueToDateProperty(Object obj, PropertyDescriptor pd, XSSFCell cell, boolean nullable,
                                            DateSourceType dateType, String pattern)
            throws DataTypeNotSupportedException, InvocationTargetException,
            IllegalAccessException, ParseException {

        String propertyTypeName = pd.getPropertyType().getName();
        Method setter = pd.getWriteMethod();

        switch (propertyTypeName){
            case ("java.util.Date"):
                Date date;
                if (dateType == DateSourceType.DATE){
                    date = cell.getDateCellValue();

                } else {
                    String rawValue = dataFormatter.formatCellValue(cell);
                    System.out.println("java.util.Date: rawValue = " + rawValue);
                    DateFormat df = new SimpleDateFormat(pattern);
                    date = df.parse(rawValue);
                }
                System.out.println("java.util.Date: date property: " + date);
                setter.invoke(obj, date);
                break;

            default:
                throw new DataTypeNotSupportedException(String.format("Type '%s' is not supported", propertyTypeName));
        }

    }

    private void setCellValueToProperty(Object obj, PropertyDescriptor pd, XSSFCell cell, boolean nullable)
            throws DataTypeNotSupportedException, InvocationTargetException, IllegalAccessException {
        String propertyTypeName = pd.getPropertyType().getName();
        Method setter = pd.getWriteMethod();

        switch (propertyTypeName){
            case ("java.lang.String"):
                String rawValue = dataFormatter.formatCellValue(cell);
                System.out.println("java.lang.String: rawValue = " + rawValue);
                setter.invoke(obj, rawValue);
                break;

            default:
                throw new DataTypeNotSupportedException(String.format("Type '%s' is not supported", propertyTypeName));
        }
    }

    private PropertyDescriptor getPropertyDescriptor(Method getter) throws IntrospectionException {
        Class<?> clazz = getter.getDeclaringClass();
        BeanInfo info = Introspector.getBeanInfo(clazz);
        PropertyDescriptor[] props = info.getPropertyDescriptors();
        for (PropertyDescriptor pd : props) {
            if (getter.equals(pd.getReadMethod())) {
                return pd;
            }
        }
        return null;
    }

    public <T extends Object> T readRow(XSSFRow excelRow, Class<T> clazz){
        return null;
    }
}
