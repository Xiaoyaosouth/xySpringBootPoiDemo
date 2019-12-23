package com.fe.mysbpoidemo.model;

import java.util.ArrayList;
import java.util.List;

/**
 * @author zky
 */
public class A {

    String[][] xlsx = new String[10][10];
    Abean[] a= new Abean[]{};
    //xlsx[0] = a;

    List<List<Abean>> sheet = new ArrayList<>();

    class Abean{
        int x;
        int y;
        String value;
        String style;
        String type;
    }


}
