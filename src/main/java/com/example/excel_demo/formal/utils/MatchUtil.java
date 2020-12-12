package com.example.excel_demo.formal.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Classname MatchUtil
 * @Description TODO
 * @Date 2020/12/10 9:54
 * @Author by ZhangLei
 */
@Slf4j
public class MatchUtil {

    /**
     * 颜色匹配
     *
     * @param cellColor
     * @param color
     * @return
     */
    public static Boolean colorMatch(String cellColor, String color) {
        if (StringUtils.isBlank(color)) {
            return false;
        }
        return StringUtils.equals(cellColor, color);
    }

    /**
     * 字体颜色匹配
     * @param fontColor
     * @param color
     * @return
     */
    public static Boolean fontColorMatch(String fontColor, String color) {
        if (StringUtils.isBlank(color)) {
            return false;
        }
        return StringUtils.equals(fontColor, color);
    }

    /**
     * 进行单元格值与定义的正则表达式的匹配
     *
     * @param cellValue
     * @param patterns
     * @return
     */
    public static Boolean patternMatch(String cellValue, String[] patterns) {
        if (patterns == null || patterns.length == 0) {
            return false;
        }
        for (String pattern : patterns) {
            Pattern p = Pattern.compile(pattern);
            Matcher matcher = p.matcher(cellValue);
            if (!matcher.find()) {
                return false;
            }
        }
        return true;
    }

    /**
     * 进行单元格的样式匹配
     * @param cell
     * @param styles
     * @return
     */
    public static Boolean styleMatch(Cell cell, Cell nextCell, String[] styles) {
        if (styles == null || styles.length == 0) {
            return false;
        }
        CellStyle cellStyle = cell.getCellStyle();
        CellStyle nextCellStyle = nextCell.getCellStyle();
        return cellStyle.getBorderBottomEnum().name().equals(styles[0])
                && nextCellStyle.getBorderBottomEnum().name().equals(styles[1]);
    }

}
