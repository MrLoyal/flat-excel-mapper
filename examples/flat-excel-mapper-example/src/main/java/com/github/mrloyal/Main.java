package com.github.mrloyal;

import com.github.mrloyal.flatexcelmapper.FlatExcelMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

public class Main {

    private static final Logger logger = LoggerFactory.getLogger(Main.class);

    public static void main(String[] args) throws Exception {
        String inputPath  = args.length > 0 ? args[0] : "files/students.xlsx";
        String outputPath = args.length > 1 ? args[1] : "files/students_output.xlsx";

        FlatExcelMapper mapper = new FlatExcelMapper();

        // --- Read from an existing file ---
        logger.info("Reading students from: {}", inputPath);
        List<Student> students = mapper.read(inputPath, 0, Student.class);
        logger.info("Found {} student(s):", students.size());
        for (Student s : students) {
            logger.info("  {}", s);
        }

        // --- Build a list of new students to write ---
        Student alice = new Student();
        alice.setOrder(10);
        alice.setName("Alice");
        alice.setBirthDate("01/01/1990");
        alice.setJoinDate(LocalDate.of(2023, 9, 1));
        alice.setGood(true);

        Student bob = new Student();
        bob.setOrder(11);
        bob.setName("Bob");
        bob.setBirthDate("15/06/1985");
        bob.setJoinDate(LocalDate.of(2024, 1, 15));
        bob.setGood(false);

        List<Student> newStudents = Arrays.asList(alice, bob);

        // --- Write to a new file ---
        logger.info("Writing {} student(s) to: {}", newStudents.size(), outputPath);
        mapper.write(newStudents, Paths.get(outputPath));
        logger.info("Write complete.");

        // --- Read back to verify ---
        logger.info("Reading back from: {}", outputPath);
        List<Student> written = mapper.read(outputPath, 0, Student.class);
        logger.info("Read back {} student(s):", written.size());
        for (Student s : written) {
            logger.info("  {}", s);
        }
    }
}
