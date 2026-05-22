package com.github.mrloyal;

import com.github.mrloyal.flatexcelmapper.FlatExcelMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

public class Main {

    private static final Logger logger = LoggerFactory.getLogger(Main.class);

    public static void main(String[] args) throws Exception {
        String filePath = args.length > 0 ? args[0] : "files/students.xlsx";

        FlatExcelMapper mapper = new FlatExcelMapper();

        logger.info("Reading all students from: {}", filePath);
        List<Student> students = mapper.read(filePath, 0, Student.class);

        logger.info("Found {} student(s):", students.size());
        for (Student s : students) {
            logger.info("  {}", s);
        }
    }
}
