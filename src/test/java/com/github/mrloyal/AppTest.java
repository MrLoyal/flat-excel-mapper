package com.github.mrloyal;

import com.github.mrloyal.flatexcelmapper.FlatExcelMapper;
import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * Unit test for simple App.
 */
public class AppTest extends TestCase {
    public static void main(String[] args) {
        FlatExcelMapper mapper = new FlatExcelMapper();

        int sheetIndex = 0;
        List<Student> students = null;
        File pwd = new File(".");
        System.out.println("PWD: " + pwd.getAbsolutePath());
        try {
            students = mapper.read("FlatExcelMapper/files/students.xlsx", sheetIndex, Student.class);
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("AppTest::main(): students: " + students);
    }
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public AppTest(String testName) {
        super(testName);
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite() {
        return new TestSuite(AppTest.class);
    }

    /**
     * Rigourous Test :-)
     */
    public void testApp() {
        assertTrue(true);
    }
}
