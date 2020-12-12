package com.example.excel_demo.formal.config;

import lombok.Data;

/**
 * @Classname Header
 * @Description TODO
 * @Date 2020/12/9 19:01
 * @Author by ZhangLei
 */
@Data
public class Header {

    private Boolean accurate;

    private String bgColor;

    private String[] pattern;

    private String[] style;

    private String fontColor;

    private Boolean needParse;

    private Integer rowRange;

}
