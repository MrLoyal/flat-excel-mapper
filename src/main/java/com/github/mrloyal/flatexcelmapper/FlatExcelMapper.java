package com.github.mrloyal.flatexcelmapper;

import com.github.mrloyal.flatexcelmapper.annotation.DateSourceType;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelColumn;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelDate;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelEntity;
import com.github.mrloyal.flatexcelmapper.exception.DataTypeNotSupportedException;
import com.github.mrloyal.flatexcelmapper.exception.EmptyCellException;
import com.github.mrloyal.flatexcelmapper.exception.ExcelMapperException;
import com.github.mrloyal.flatexcelmapper.exception.InvalidValueException;
import org.apache.poi.hssf.util.CellReference;
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
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
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
            throws EmptyCellException, DataTypeNotSupportedException, InvalidValueException, ExcelMapperException{
        List<T> list = new ArrayList<T>();
        XSSFWorkbook workbook;
        FileInputStream fis = null;
        IOException ioException = null;
        XSSFSheet sheet;
        int dataStartRow = 1;
        if (clazz.isAnnotationPresent(ExcelEntity.class)){
            ExcelEntity entityAnn = clazz.getAnnotation(ExcelEntity.class);
            dataStartRow = entityAnn.dataStartRow();
        }

        // System.out.println("Data start row = " + dataStartRow);
        try {
            fis = new FileInputStream(file);
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(sheetIndex);

            XSSFRow startRow = sheet.getRow(dataStartRow - 1);

            Iterator<Row> iterator = sheet.rowIterator();
            while (iterator.hasNext()){
                XSSFRow row1 = (XSSFRow) iterator.next();
                if (row1 != null && row1.equals(startRow)){
                    // System.out.println(row1.getCell(1));
                    // System.out.println("Ok, got him");
                    T obj = readRow(row1, clazz);
                    list.add(obj);
                    break;
                }
            }

            while (iterator.hasNext()){
                XSSFRow row = (XSSFRow) iterator.next();
                //XSSFCell cell = row.getCell(1);
                //System.out.println("cell: " + cell);
                T obj = readRow(row, clazz);
                list.add(obj);
            }

            // readRow(sheet, 5, clazz);
        } catch (IOException e){
            ioException = e;
        } finally {
            if (fis != null){
                try {
                    fis.close();
                } catch (IOException e) {
                    ioException = e;
                }
            }
        }

        if (ioException != null){
            throw new ExcelMapperException(ioException);
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

    private void setCellValueToDateProperty(Object obj, PropertyDescriptor pd, XSSFCell cell, boolean nullable,
                                            DateSourceType dateType, String pattern)
            throws DataTypeNotSupportedException, InvocationTargetException,
            IllegalAccessException, ParseException {

        String propertyTypeName = pd.getPropertyType().getName();
        Method setter = pd.getWriteMethod();
        String rawValue = dataFormatter.formatCellValue(cell);
        switch (propertyTypeName){
            case ("java.util.Date"):
                Date date;
                if (dateType == DateSourceType.DATE){
                    date = cell.getDateCellValue();

                } else {
                    // System.out.println("java.util.Date: rawValue = " + rawValue);
                    DateFormat df = new SimpleDateFormat(pattern);
                    date = df.parse(rawValue);
                }
                // System.out.println("java.util.Date: date property: " + date);
                setter.invoke(obj, date);
                break;

            case ("java.time.LocalDate"):
                LocalDate localDate;
                if (dateType == DateSourceType.DATE){
                    date = cell.getDateCellValue();
                    localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                } else {
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(pattern);
                    localDate = LocalDate.parse(rawValue, formatter);
                }
                // System.out.println("java.time.LocalDate: local date property: " + localDate);
                setter.invoke(obj, localDate);
                break;
            default:
                throw new DataTypeNotSupportedException(String.format("Type '%s' is not supported", propertyTypeName));
        }

    }

    private void setCellValueToProperty(Object obj, PropertyDescriptor pd, XSSFCell cell, boolean nullable)
            throws EmptyCellException, DataTypeNotSupportedException, ExcelMapperException, InvalidValueException {
        String propertyTypeName = pd.getPropertyType().getName();
        Method setter = pd.getWriteMethod();

        try{
            String rawValue = dataFormatter.formatCellValue(cell);
            // Cannot be empty but got an empty cell
            if (!nullable && (rawValue == null || "".equals(rawValue.trim()))){
                throwEmptyCellException(cell);
            }
            switch (propertyTypeName){
                case ("int"):
                    int intVal;
                    // Property marked as nullable and got an empty cell
                    if (rawValue == null || "".equals(rawValue.trim())){
                        // Set value of this property to default value
                        intVal = 0;
                        setter.invoke(obj, intVal);
                    }

                    // Property marked as nullable and got cell with value
                    else {

                        // Attempt to parse cell value into desired type
                        try{
                            intVal = Integer.valueOf(rawValue);
                            setter.invoke(obj, intVal);
                        } catch (NumberFormatException nfe){
                            throwInvalidValueException(nfe, cell);
                        }

                    }
                    break;

                case ("boolean"):
                    boolean boolVal;
                    if (rawValue == null || "".equals(rawValue.trim())){
                        boolVal = false;
                        setter.invoke(obj, boolVal);
                    }
                    else {
                        try{
                            boolVal = Boolean.parseBoolean(rawValue);
                            setter.invoke(obj, boolVal);
                        } catch (Exception nfe){
                            throwInvalidValueException(nfe, cell);
                        }

                    }
                    break;

                case ("byte"):
                    byte byteVal;
                    if (rawValue == null || "".equals(rawValue.trim())){
                        byteVal = 0;
                        setter.invoke(obj, byteVal);
                    }
                    else {
                        try{
                            byteVal = Byte.parseByte(rawValue);
                            setter.invoke(obj, byteVal);
                        } catch (NumberFormatException nfe){
                            throwInvalidValueException(nfe, cell);
                        }

                    }
                    break;

                case ("java.lang.String"):

                    if (!nullable && (rawValue == null || "".equals(rawValue.trim()))){
                        throwEmptyCellException(cell);
                    } else {
                        setter.invoke(obj, rawValue);
                    }
                    break;

                case (""):
                    break;

                default:
                    throw new DataTypeNotSupportedException(String.format("Type '%s' is not supported", propertyTypeName));
            }
        } catch (IllegalAccessException e) {
            throw new ExcelMapperException(e);
        } catch (InvocationTargetException e) {
            throw new ExcelMapperException(e);
        }

    }

    private void throwEmptyCellException(XSSFCell cell) throws EmptyCellException{
        String msg;
        if (cell != null){
            msg = String.format("Cell is empty but value is required. Cell (row = %d, col = %d)",
                    cell.getAddress().getRow(), cell.getAddress().getColumn());
        } else {
            msg = "Cell is empty but value is required.";
        }

        throw new EmptyCellException(msg, cell);
    }

    private void throwInvalidValueException(Exception ex, XSSFCell cell) throws InvalidValueException{
        String msg = String.format("Value is invalid. Cell (row = %d, col = %d)",
                cell.getAddress().getRow(), cell.getAddress().getColumn());
        throw new InvalidValueException(msg, ex, cell);
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


    public <T extends Object> T readRow(XSSFSheet sheet, int humanReadableRow, Class<T> clazz)
            throws EmptyCellException, DataTypeNotSupportedException, ExcelMapperException, InvalidValueException {

        XSSFRow row = sheet.getRow(humanReadableRow-1);
        return readRow(row, clazz);
    }

    public <T extends Object> T readRow(XSSFRow excelRow, Class<T> clazz)
            throws EmptyCellException, DataTypeNotSupportedException, ExcelMapperException, InvalidValueException {
        T object;
        try {
            object = clazz.newInstance();
            XSSFCell cell;
            for (Method method : clazz.getDeclaredMethods()){
                if (method.isAnnotationPresent(ExcelColumn.class)){
                    ExcelColumn colAnn = method.getAnnotation(ExcelColumn.class);
                    boolean nullable = colAnn.nullable();
                    String colName = colAnn.name();
                    int colIndex = CellReference.convertColStringToIndex(colName);
                    cell = excelRow.getCell(colIndex);
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
        } catch (InstantiationException e) {
            throw new ExcelMapperException(e);
        } catch (InvocationTargetException e) {
            throw new ExcelMapperException(e);
        } catch (IntrospectionException e) {
            throw new ExcelMapperException(e);
        } catch (IllegalAccessException e) {
            throw new ExcelMapperException(e);
        } catch (ParseException e) {
            throw new ExcelMapperException(e);
        } 

        return object;
    }
}
