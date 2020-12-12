package com.example.excel_demo.formal.bean;

import lombok.Data;

import java.util.List;

/**
 * @Classname VslVoyList
 * @Description TODO
 * @Date 2020/12/11 15:50
 * @Author by ZhangLei
 */
@Data
public class VslVoys {

    /**
     * 船期
     */
    private List<VslVoy> vslVoyList;

    /**
     * 备注信息
     */
    private String remarks;

}
