package com.example.excel_demo.formal.config;

import lombok.Data;

/**
 * @Classname Content
 * @Description TODO
 * @Date 2020/12/9 19:01
 * @Author by ZhangLei
 */
@Data
public class Content {

    private TitleFeature titleFeature;

    private String[] column;

    private String[] special;

    private String[] ignoreColumn;

}
