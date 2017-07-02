package com.github.mrloyal.flatexcelmapper.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.METHOD)
public @interface ExcelColumn {
    String name();
    boolean nullable() default false;

    DateSourceType dateSourceType() default DateSourceType.DATE;
    String datePattern() default "yyyy/MM/dd";

}