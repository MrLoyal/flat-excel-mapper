package com.github.mrloyal;

import com.github.mrloyal.flatexcelmapper.FlatExcelMapper;
import com.github.mrloyal.flatexcelmapper.exception.ExcelMapperException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.GregorianCalendar;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class FlatExcelMapperWriteTest {

    private FlatExcelMapper mapper;
    private List<Student> students;

    @BeforeEach
    void setUp() {
        mapper = new FlatExcelMapper();
        students = buildStudents();
    }

    private List<Student> buildStudents() {
        Student s1 = new Student();
        s1.setOrder(1);
        s1.setName("Lorem");
        s1.setBirthDate(new GregorianCalendar(2017, Calendar.JULY, 18).getTime());
        s1.setJoinDate(LocalDate.of(2020, 3, 15));
        s1.setGood(true);

        Student s2 = new Student();
        s2.setOrder(2);
        s2.setName("Ipsum");
        s2.setBirthDate(new GregorianCalendar(1995, Calendar.DECEMBER, 25).getTime());
        s2.setJoinDate(LocalDate.of(2021, 6, 1));
        s2.setGood(false);

        return Arrays.asList(s1, s2);
    }

    @Test
    void writeToSheetCreatesCorrectNumberOfRows() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        mapper.write(students, sheet);
        // Student.dataStartRow = 5, so two students occupy rows 4 and 5 (0-indexed)
        assertEquals(2, sheet.getPhysicalNumberOfRows());
        workbook.close();
    }

    @Test
    void writeToSheetRoundTrip() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        mapper.write(students, sheet);

        Path tmp = Files.createTempFile("write_sheet_test", ".xlsx");
        try {
            try (FileOutputStream fos = new FileOutputStream(tmp.toFile())) {
                workbook.write(fos);
            }
            workbook.close();

            List<Student> result = mapper.read(tmp.toFile(), 0, Student.class);
            assertEquals(2, result.size());
            assertStudent1(result.get(0));
            assertStudent2(result.get(1));
        } finally {
            Files.deleteIfExists(tmp);
        }
    }

    @Test
    void writeToPathRoundTrip() throws Exception {
        Path tmp = Files.createTempFile("write_path_test", ".xlsx");
        try {
            mapper.write(students, tmp);

            List<Student> result = mapper.read(tmp.toFile(), 0, Student.class);
            assertEquals(2, result.size());
            assertStudent1(result.get(0));
            assertStudent2(result.get(1));
        } finally {
            Files.deleteIfExists(tmp);
        }
    }

    @Test
    void writeToPathPreservesIntField() throws Exception {
        Path tmp = Files.createTempFile("write_int_test", ".xlsx");
        try {
            mapper.write(students, tmp);
            List<Student> result = mapper.read(tmp.toFile(), 0, Student.class);
            assertEquals(1, result.get(0).getOrder());
            assertEquals(2, result.get(1).getOrder());
        } finally {
            Files.deleteIfExists(tmp);
        }
    }

    @Test
    void writeToPathPreservesStringField() throws Exception {
        Path tmp = Files.createTempFile("write_string_test", ".xlsx");
        try {
            mapper.write(students, tmp);
            List<Student> result = mapper.read(tmp.toFile(), 0, Student.class);
            assertEquals("Lorem", result.get(0).getName());
            assertEquals("Ipsum", result.get(1).getName());
        } finally {
            Files.deleteIfExists(tmp);
        }
    }

    @Test
    void writeToPathPreservesBooleanField() throws Exception {
        Path tmp = Files.createTempFile("write_bool_test", ".xlsx");
        try {
            mapper.write(students, tmp);
            List<Student> result = mapper.read(tmp.toFile(), 0, Student.class);
            assertTrue(result.get(0).isGood());
            assertFalse(result.get(1).isGood());
        } finally {
            Files.deleteIfExists(tmp);
        }
    }

    @Test
    void writeToPathPreservesDateStringField() throws Exception {
        Path tmp = Files.createTempFile("write_date_string_test", ".xlsx");
        try {
            mapper.write(students, tmp);
            List<Student> result = mapper.read(tmp.toFile(), 0, Student.class);

            Calendar cal = Calendar.getInstance();

            cal.setTime(result.get(0).getBirthDate());
            assertEquals(2017, cal.get(Calendar.YEAR));
            assertEquals(Calendar.JULY, cal.get(Calendar.MONTH));
            assertEquals(18, cal.get(Calendar.DAY_OF_MONTH));

            cal.setTime(result.get(1).getBirthDate());
            assertEquals(1995, cal.get(Calendar.YEAR));
            assertEquals(Calendar.DECEMBER, cal.get(Calendar.MONTH));
            assertEquals(25, cal.get(Calendar.DAY_OF_MONTH));
        } finally {
            Files.deleteIfExists(tmp);
        }
    }

    @Test
    void writeToPathPreservesLocalDateField() throws Exception {
        Path tmp = Files.createTempFile("write_localdate_test", ".xlsx");
        try {
            mapper.write(students, tmp);
            List<Student> result = mapper.read(tmp.toFile(), 0, Student.class);
            assertEquals(LocalDate.of(2020, 3, 15), result.get(0).getJoinDate());
            assertEquals(LocalDate.of(2021, 6, 1), result.get(1).getJoinDate());
        } finally {
            Files.deleteIfExists(tmp);
        }
    }

    @Test
    void writeEmptyListDoesNothing() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        mapper.write(Collections.emptyList(), sheet);
        assertEquals(0, sheet.getPhysicalNumberOfRows());
        workbook.close();
    }

    @Test
    void writeThrowsForClassWithoutExcelEntityAnnotation() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        assertThrows(ExcelMapperException.class,
                () -> mapper.write(Collections.singletonList(new UnannotatedEntity()), sheet));
    }

    private void assertStudent1(Student s) {
        assertEquals(1, s.getOrder());
        assertEquals("Lorem", s.getName());
        assertTrue(s.isGood());
        assertEquals(LocalDate.of(2020, 3, 15), s.getJoinDate());
        Calendar cal = Calendar.getInstance();
        cal.setTime(s.getBirthDate());
        assertEquals(2017, cal.get(Calendar.YEAR));
        assertEquals(Calendar.JULY, cal.get(Calendar.MONTH));
        assertEquals(18, cal.get(Calendar.DAY_OF_MONTH));
    }

    private void assertStudent2(Student s) {
        assertEquals(2, s.getOrder());
        assertEquals("Ipsum", s.getName());
        assertFalse(s.isGood());
        assertEquals(LocalDate.of(2021, 6, 1), s.getJoinDate());
        Calendar cal = Calendar.getInstance();
        cal.setTime(s.getBirthDate());
        assertEquals(1995, cal.get(Calendar.YEAR));
        assertEquals(Calendar.DECEMBER, cal.get(Calendar.MONTH));
        assertEquals(25, cal.get(Calendar.DAY_OF_MONTH));
    }

    static class UnannotatedEntity {
        public String getValue() { return "x"; }
    }
}
