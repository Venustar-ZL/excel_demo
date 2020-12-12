package com.example.excel_demo.formal.utils;

import com.alibaba.fastjson.JSONObject;
import com.example.excel_demo.formal.config.ConfigBean;
import com.example.excel_demo.formal.config.VslVoyAttribute;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @Classname YamlUtil
 * @Description TODO
 * @Date 2020/12/9 18:39
 * @Author by ZhangLei
 */
@Slf4j
public class YamlUtil {

    public static void main(String[] args) {
        ConfigBean configBean = getConfigBean("application-Test");
        System.out.println(JSONObject.toJSONString(configBean));
    }

    /**
     * 通过配置文件名称获取配置文件
     * @param yamlName
     * @return
     */
    public static ConfigBean getConfigBean(String yamlName) {
        ObjectMapper mapper = new ObjectMapper(new YAMLFactory());
        String resourceName = "/" + yamlName + ".yml";
        String path = YamlUtil.class.getResource(resourceName).getFile();
        ConfigBean configBean = null;
        try {
            configBean = mapper.readValue(new File(path), ConfigBean.class);
        } catch (IOException e) {
            e.printStackTrace();
            log.info("读取配置文件[{}]失败", yamlName);
        }
        return configBean;
    }

    /**
     * 根据航线名称获取配置文件
     * @param lineName
     * @return
     */
    public static ConfigBean getConfigBeanByLine(String lineName) {
        ObjectMapper mapper = new ObjectMapper(new YAMLFactory());
        String resourceName = "/" + lineName + ".yml";
        String path = YamlUtil.class.getResource(resourceName).getFile();
        ConfigBean configBean = null;
        try {
            configBean = mapper.readValue(new File(path), ConfigBean.class);
        } catch (IOException e) {
            e.printStackTrace();
            log.info("读取配置文件[{}]失败", lineName);
        }
        return configBean;
    }


    /**
     * 解析通用属性配置
     */
    public static List<VslVoyAttribute> parseUniversal(ConfigBean configBean) {
        List<VslVoyAttribute> attributeList = new ArrayList<>();
        String[] universalList = configBean.getContent().getColumn();
        for (String universal : universalList) {
            VslVoyAttribute vslVoyAttribute = new VslVoyAttribute();
            vslVoyAttribute.setName(universal.substring(0, universal.indexOf("(")));
            vslVoyAttribute.setLength(Integer.parseInt(universal.substring(universal.indexOf("(") + 1, universal.indexOf(")"))));
            attributeList.add(vslVoyAttribute);
        }
        return attributeList;
    }

    /**
     * 解析特殊属性配置
     */
    public static VslVoyAttribute parseSpecial(ConfigBean configBean) {
        String[] special = configBean.getContent().getSpecial();
        VslVoyAttribute vslVoyAttribute = new VslVoyAttribute();
        vslVoyAttribute.setBegin(special[0]);
        vslVoyAttribute.setEnd(special[1]);
        return vslVoyAttribute;
    }

}
