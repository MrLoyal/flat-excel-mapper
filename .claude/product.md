# flat-excel-mapper

## What it is

A Java library that maps rows from `.xlsx` Excel files to POJOs using annotations — similar in spirit to JPA/Hibernate ORM, but for spreadsheets instead of databases.

## How it works

You annotate a plain Java class with:

- `@ExcelEntity(dataStartRow = N)` — marks the class as mappable and tells the mapper which row to start reading data from (useful when the spreadsheet has headers or metadata rows above the data).
- `@ExcelColumn(name = "B")` on a getter — maps that field to a named Excel column (A, B, C, …).
- `@ExcelDate(type = DateSourceType.DATE|STRING, format = "dd/MM/yyyy")` — additionally placed on date getters to handle `java.util.Date` and `java.time.LocalDate` fields, reading either from a native Excel date cell or a string-formatted cell.
- `nullable = true/false` on `@ExcelColumn` — controls whether an empty cell throws `EmptyCellException`.

The `FlatExcelMapper` class is the entry point. Its `read(file, sheetIndex, MyClass.class)` method returns a `List<MyClass>` with one instance per data row.

## Supported field types

- `int`, `byte`, `boolean` (Java primitives)
- `java.lang.String`
- `java.util.Date`
- `java.time.LocalDate`

## Tech stack

- Java 8, Maven
- Apache POI 4.1.1 (XSSF — `.xlsx` only)
- JUnit 5 (Jupiter + Vintage)

## Current state

Core read path is implemented and tested. Stubs exist for reading from an already-open `XSSFWorkbook` or `XSSFSheet` but return `null` — not yet implemented.

## Example usage

```java
@ExcelEntity(dataStartRow = 5)
public class Student {
    private String name;
    private Date birthDate;
    private LocalDate joinDate;

    @ExcelColumn(name = "B")
    public String getName() { return name; }

    @ExcelColumn(name = "C")
    @ExcelDate(type = DateSourceType.STRING, format = "dd/MM/yyyy")
    public Date getBirthDate() { return birthDate; }

    @ExcelColumn(name = "D")
    @ExcelDate(type = DateSourceType.DATE)
    public LocalDate getJoinDate() { return joinDate; }
    // setters ...
}

List<Student> students = new FlatExcelMapper().read("students.xlsx", 0, Student.class);
```
