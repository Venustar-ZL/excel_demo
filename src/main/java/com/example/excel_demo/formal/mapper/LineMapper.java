package com.example.excel_demo.formal.mapper;

import lombok.extern.slf4j.Slf4j;

import javax.annotation.PostConstruct;
import java.util.HashMap;
import java.util.Map;

/**
 * @Classname LineMapper
 * @Description TODO
 * @Date 2020/12/11 10:02
 * @Author by ZhangLei
 */
@Slf4j
public class LineMapper {

    private static final Map<String, String> LINE_MAPPER = new HashMap<>();

//    @PostConstruct
    public static void init() {
//        log.info("加载航线映射表ing...........");
        LINE_MAPPER.put("AAC2", "application-MJ");
        LINE_MAPPER.put("AACI", "application-MJ");
        LINE_MAPPER.put("CEN", "application-MJ");
        LINE_MAPPER.put("CPNW", "application-MJ");
        LINE_MAPPER.put("MPNW", "application-MJ");
        LINE_MAPPER.put("EPNW", "application-MJ");
        LINE_MAPPER.put("AWE1", "application-MJ");
        LINE_MAPPER.put("AWE2", "application-MJ");
        LINE_MAPPER.put("AWE4", "application-MJ");
        LINE_MAPPER.put("GME2", "application-MJ");
        LINE_MAPPER.put("GME", "application-MJ");
        LINE_MAPPER.put("AAC4", "application-MJ");

        LINE_MAPPER.put("ESA", "application-LF");
        LINE_MAPPER.put("ESA2", "application-LF");
        LINE_MAPPER.put("WSA", "application-LF");
        LINE_MAPPER.put("WSA2", "application-LF");
        LINE_MAPPER.put("WSA3", "application-LF");
        LINE_MAPPER.put("WSA4", "application-LF");
        LINE_MAPPER.put("CAX1", "application-LF");
        LINE_MAPPER.put("ZAX1", "application-LF");
        LINE_MAPPER.put("ZAX3", "application-LF");
        LINE_MAPPER.put("WAX1", "application-LF");
        LINE_MAPPER.put("WAX3", "application-LF");
        LINE_MAPPER.put("WAX4", "application-LF");
        LINE_MAPPER.put("EAX1", "application-LF");
        LINE_MAPPER.put("EAX3", "application-LF");

        LINE_MAPPER.put("PA1", "application-DNY");
        LINE_MAPPER.put("CKI", "application-DNY");
        LINE_MAPPER.put("CTI1", "application-DNY");
        LINE_MAPPER.put("CT2", "application-DNY");
        LINE_MAPPER.put("CT3", "application-DNY");
        LINE_MAPPER.put("KCM2", "application-DNY");
        LINE_MAPPER.put("RBC2", "application-DNY");
        LINE_MAPPER.put("CMS2", "application-DNY");
        LINE_MAPPER.put("CV1", "application-DNY");
        LINE_MAPPER.put("CV2", "application-DNY");
        LINE_MAPPER.put("CHS", "application-DNY");
        LINE_MAPPER.put("CV5", "application-DNY");
        LINE_MAPPER.put("CNP2", "application-DNY");
        LINE_MAPPER.put("CPF", "application-DNY");
        LINE_MAPPER.put("CIX3", "application-DNY");
        LINE_MAPPER.put("PMX", "application-DNY");
        LINE_MAPPER.put("FCE", "application-DNY");
        LINE_MAPPER.put("AS1", "application-DNY");
        LINE_MAPPER.put("CPX", "application-DNY");
        LINE_MAPPER.put("AK6", "application-DNY");
        LINE_MAPPER.put("AK12", "application-DNY");
        LINE_MAPPER.put("AEU3", "application-DNY");
        LINE_MAPPER.put("AEU9", "application-DNY");
//        log.info("加载航线映射表完毕！");
    }

    public static Map<String, String> getLineMapper() {
        return LINE_MAPPER;
    }

}
