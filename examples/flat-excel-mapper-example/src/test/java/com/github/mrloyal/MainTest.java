package com.github.mrloyal;

import com.github.mrloyal.flatexcelmapper.FlatExcelMapper;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("flat-excel-mapper example: read students.xlsx")
public class MainTest {

    private static final String FILE_PATH = "files/students.xlsx";
    private FlatExcelMapper mapper;

    @BeforeEach
    void setUp() {
        mapper = new FlatExcelMapper();
    }

    @Test
    @DisplayName("reads all student rows")
    void readsAllStudents() throws Exception {
        List<Student> students = mapper.read(FILE_PATH, 0, Student.class);
        assertEquals(3, students.size());
    }

    @Test
    @DisplayName("maps order and name correctly")
    void mapsOrderAndName() throws Exception {
        List<Student> students = mapper.read(FILE_PATH, 0, Student.class);

        assertEquals(1, students.get(0).getOrder());
        assertEquals("Lorem", students.get(0).getName());

        assertEquals(2, students.get(1).getOrder());
        assertEquals("Ipsum", students.get(1).getName());

        assertEquals(3, students.get(2).getOrder());
        assertEquals("Haha", students.get(2).getName());
    }

    @Test
    @DisplayName("maps birthDate as string")
    void mapsBirthDate() throws Exception {
        List<Student> students = mapper.read(FILE_PATH, 0, Student.class);
        assertEquals("18/07/2017", students.get(0).getBirthDate());
    }

    @Test
    @DisplayName("maps joinDate as LocalDate via DateSourceType.DATE")
    void mapsJoinDate() throws Exception {
        List<Student> students = mapper.read(FILE_PATH, 0, Student.class);
        assertNotNull(students.get(0).getJoinDate());
        assertNotNull(students.get(1).getJoinDate());
        assertNotNull(students.get(2).getJoinDate());
    }
}
