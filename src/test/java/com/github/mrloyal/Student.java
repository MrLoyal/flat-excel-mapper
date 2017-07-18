package com.github.mrloyal;

import com.github.mrloyal.flatexcelmapper.annotation.DateSourceType;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelColumn;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelDate;
import com.github.mrloyal.flatexcelmapper.annotation.ExcelEntity;

import java.util.Date;

/**
 * MIT License
 * <p>
 * Copyright (c) 2017 Thành Loyal
 * <p>
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * <p>
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * <p>
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

@ExcelEntity(dataStartRow = 5)
public class Student {
    private String name;
    private Date birthDate;

    @ExcelColumn(name = "B")
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @ExcelColumn(name = "C")
    @ExcelDate(type = DateSourceType.STRING, format = "dd/MM/yyyy")
    public Date getBirthDate() {
        return birthDate;
    }

    public void setBirthDate(Date birthDate) {
        this.birthDate = birthDate;
    }
}
