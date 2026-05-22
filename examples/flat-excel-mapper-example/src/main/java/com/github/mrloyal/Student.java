package com.github.mrloyal;

import com.github.mrloyal.flatexcelmapper.annotation.DateSourceType;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelColumn;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelDate;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelEntity;

import java.time.LocalDate;

@ExcelEntity(dataStartRow = 5)
public class Student {

    private int order;
    private String name;
    private String birthDate;
    private LocalDate joinDate;
    private boolean isGood;

    @ExcelColumn(name = "A", nullable = true)
    public int getOrder() {
        return order;
    }

    public void setOrder(int order) {
        this.order = order;
    }

    @ExcelColumn(name = "B")
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @ExcelColumn(name = "C")
    public String getBirthDate() {
        return birthDate;
    }

    public void setBirthDate(String birthDate) {
        this.birthDate = birthDate;
    }

    @ExcelColumn(name = "D")
    @ExcelDate(type = DateSourceType.DATE)
    public LocalDate getJoinDate() {
        return joinDate;
    }

    public void setJoinDate(LocalDate joinDate) {
        this.joinDate = joinDate;
    }

    @ExcelColumn(name = "E")
    public boolean isGood() {
        return isGood;
    }

    public void setGood(boolean good) {
        isGood = good;
    }

    @Override
    public String toString() {
        return "Student{order=" + order
                + ", name='" + name + "'"
                + ", birthDate='" + birthDate + "'"
                + ", joinDate=" + joinDate
                + ", isGood=" + isGood + "}";
    }
}
