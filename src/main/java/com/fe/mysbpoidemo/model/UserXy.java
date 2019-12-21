package com.fe.mysbpoidemo.model;

import lombok.Data;

@Data
public class UserXy {
    private String name;
    private String age;
    private String phone;
    private String email;

    @Override
    public String toString() {
        return "UserXy{" +
                "name='" + name + '\'' +
                "age='" + age + '\'' +
                "phone='" + phone + '\'' +
                "email='" + email + '\'' +
                '}';
    }
}
