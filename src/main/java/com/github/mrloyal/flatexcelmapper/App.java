package com.github.mrloyal.flatexcelmapper;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) {
        App app = new App();
        Date d1 = app.test(Date.class);
        System.out.println(d1);
        List<Date> lst = app.testList(Date.class);
        System.out.println(lst);
    }

    public <T extends Object> T test(Class<T> clazz){
        T t = null;
        try {
            t = clazz.newInstance();
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return t;
    }

    public <T extends Object>List<T> testList(Class<T> clazz){
        List<T> list = new ArrayList<T>();
        T t = null;
        try {
            t = clazz.newInstance();
            list.add(t);
            t = clazz.newInstance();
            list.add(t);
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }

        return list;
    }
}
