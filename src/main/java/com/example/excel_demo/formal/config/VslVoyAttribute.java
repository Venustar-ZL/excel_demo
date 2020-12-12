package com.example.excel_demo.formal.config;

import lombok.Data;

/**
 * @Classname VslVoyAttribute
 * @Description TODO
 * @Date 2020/12/10 9:12
 * @Author by ZhangLei
 */
@Data
public class VslVoyAttribute {

    private String name;

    private Integer length;

    /**
     * 圈定可变属性的范围
     */
    private String begin;

    private String end;

}
