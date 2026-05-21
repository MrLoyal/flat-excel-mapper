package com.github.mrloyal;

import com.github.mrloyal.flatexcelmapper.FlatExcelMapper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class FlatExcelMapperReadTest {

    private File testFile;
    private FlatExcelMapper mapper;

    @BeforeEach
    void setUp() throws Exception {
        mapper = new FlatExcelMapper();
        testFile = buildTestExcelFile();
    }

    @AfterEach
    void tearDown() {
        if (testFile != null) {
            testFile.delete();
        }
    }

    // Student has @ExcelEntity(dataStartRow = 5), so rows 1-4 are skipped.
    // Column mapping: A=order, B=name, C=birthDate(string "dd/MM/yyyy"), D=joinDate(date cell), E=isGood
    private File buildTestExcelFile() throws Exception {
        File file = File.createTempFile("students_test", ".xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(workbook.createDataFormat().getFormat("dd/MM/yyyy"));

        // Row 5 (0-indexed: 4) — Student 1
        XSSFRow row1 = sheet.createRow(4);
        row1.createCell(0).setCellValue(1);
        row1.createCell(1).setCellValue("Lorem");
        row1.createCell(2).setCellValue("18/07/2017");
        XSSFCell joinDate1 = row1.createCell(3);
        joinDate1.setCellValue(new GregorianCalendar(2020, Calendar.MARCH, 15).getTime());
        joinDate1.setCellStyle(dateStyle);
        row1.createCell(4).setCellValue("true");

        // Row 6 (0-indexed: 5) — Student 2
        XSSFRow row2 = sheet.createRow(5);
        row2.createCell(0).setCellValue(2);
        row2.createCell(1).setCellValue("Ipsum");
        row2.createCell(2).setCellValue("25/12/1995");
        XSSFCell joinDate2 = row2.createCell(3);
        joinDate2.setCellValue(new GregorianCalendar(2021, Calendar.JUNE, 1).getTime());
        joinDate2.setCellStyle(dateStyle);
        row2.createCell(4).setCellValue("false");

        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        }
        workbook.close();
        return file;
    }

    @Test
    void readReturnsCorrectCount() throws Exception {
        List<Student> students = mapper.read(testFile, 0, Student.class);
        assertEquals(2, students.size());
    }

    @Test
    void readMapsOrderCorrectly() throws Exception {
        List<Student> students = mapper.read(testFile, 0, Student.class);
        assertEquals(1, students.get(0).getOrder());
        assertEquals(2, students.get(1).getOrder());
    }

    @Test
    void readMapsNameCorrectly() throws Exception {
        List<Student> students = mapper.read(testFile, 0, Student.class);
        assertEquals("Lorem", students.get(0).getName());
        assertEquals("Ipsum", students.get(1).getName());
    }

    @Test
    void readMapsBirthDateCorrectly() throws Exception {
        List<Student> students = mapper.read(testFile, 0, Student.class);

        Calendar cal = Calendar.getInstance();

        cal.setTime(students.get(0).getBirthDate());
        assertEquals(2017, cal.get(Calendar.YEAR));
        assertEquals(Calendar.JULY, cal.get(Calendar.MONTH));
        assertEquals(18, cal.get(Calendar.DAY_OF_MONTH));

        cal.setTime(students.get(1).getBirthDate());
        assertEquals(1995, cal.get(Calendar.YEAR));
        assertEquals(Calendar.DECEMBER, cal.get(Calendar.MONTH));
        assertEquals(25, cal.get(Calendar.DAY_OF_MONTH));
    }

    @Test
    void readMapsJoinDateCorrectly() throws Exception {
        List<Student> students = mapper.read(testFile, 0, Student.class);
        assertEquals(LocalDate.of(2020, 3, 15), students.get(0).getJoinDate());
        assertEquals(LocalDate.of(2021, 6, 1), students.get(1).getJoinDate());
    }

    @Test
    void readMapsIsGoodCorrectly() throws Exception {
        List<Student> students = mapper.read(testFile, 0, Student.class);
        assertTrue(students.get(0).isGood());
        assertFalse(students.get(1).isGood());
    }
}
