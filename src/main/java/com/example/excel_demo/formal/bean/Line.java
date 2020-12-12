package com.example.excel_demo.formal.bean;

import lombok.Data;

/**
 * @Classname Line
 * @Description TODO
 * @Date 2020/12/11 15:20
 * @Author by ZhangLei
 */
@Data
public class Line {

    /**
     * 航线代码
     */
    private String lineName;

    private VslVoys vslVoys;

}
