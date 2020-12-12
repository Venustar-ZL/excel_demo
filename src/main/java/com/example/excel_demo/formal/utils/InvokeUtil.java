package com.example.excel_demo.formal.utils;

import com.example.excel_demo.formal.bean.PortOfCall;
import com.example.excel_demo.formal.bean.VslVoy;

import java.lang.reflect.Field;
import java.util.List;

/**
 * @Classname InvokeUtil
 * @Description TODO
 * @Date 2020/12/10 11:06
 * @Author by ZhangLei
 */
public class InvokeUtil {

    /**
     * 通用属性设置
     * @param vslVoy
     * @param propertyName
     * @param value
     */
    public static void dynamicSet(VslVoy vslVoy, String propertyName, Object value){
        if ("null".equals(propertyName)) {
            return;
        }
        try {
            Field field  = vslVoy.getClass().getDeclaredField(propertyName);
            field.setAccessible(true);
            field.set(vslVoy, value);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 挂靠港属性设置
     * @param vslVoy
     * @param value
     * @param portOfCallNo
     * @param portOfCallList
     */
    public static void dynamicListAdd(VslVoy vslVoy, String value, Integer portOfCallNo, List<PortOfCall> portOfCallList) {
        PortOfCall portOfCall = new PortOfCall();
        portOfCall.setEta(value);
        portOfCall.setPortOfCallNo(String.valueOf(portOfCallNo));
        portOfCall.setPortOfCallName(portOfCallList.get(portOfCallNo - 1).getPortOfCallName());
        vslVoy.getPortOfCalls().add(portOfCall);
    }

}
