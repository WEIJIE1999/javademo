package com.wj.bean;

import lombok.Data;

import java.util.Date;

/**
 * @author weijie
 * @create 2021-01-07 下午 23:34
 */
@Data
public class User {
    private String name;
    private String age;
    private Date birthDay;

    public User(String name, String age, Date birthDay) {
        this.name = name;
        this.age = age;
        this.birthDay = birthDay;
    }
}
