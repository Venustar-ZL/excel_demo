package com.example.excel_demo.formal.config;

import lombok.Data;

/**
 * @Classname Tail
 * @Description TODO
 * @Date 2020/12/9 19:01
 * @Author by ZhangLei
 */
@Data
public class Tail {

    private Boolean accurate;

    private String bgColor;

    private String[] pattern;

    private String[] style;

    private String fontColor;

    private Boolean needParse;

    private Integer rowRange;

}
