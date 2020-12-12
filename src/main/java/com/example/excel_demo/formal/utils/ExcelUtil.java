package com.example.excel_demo.formal.utils;

import com.alibaba.excel.util.DateUtils;
import com.example.excel_demo.formal.bean.PortOfCall;
import com.example.excel_demo.formal.config.*;
import com.example.excel_demo.formal.config.Header;
import com.example.excel_demo.formal.mapper.LineMapper;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.text.DecimalFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Classname ExcelUtil
 * @Description TODO
 * @Date 2020/12/9 18:00
 * @Author by ZhangLei
 */
@Slf4j
public class ExcelUtil {

    public static void main(String[] args) {
        String header = "上海/东南亚线（PA1）";
        String res = getLineNameInHeader(header);
        System.out.println(res);
    }

    /**
     * 获取Excel文件类型
     * @param file
     * @return
     */
    public static String getExcelType(File file) {
        String type = "";
        String fileName = file.getName();
        if (fileName.endsWith("xls")) {
            type = "xls";
        }
        if (fileName.endsWith("xlsx")) {
            type = "xlsx";
        }
        return type;
    }


    /**
     * 获取所有不隐藏的sheet
     * @param file
     * @param type
     * @return
     */
    public static Iterator<Sheet> getAllNotHiddenSheet(File file, String type) {
        List<Sheet> sheetList = new ArrayList<>();
        Workbook workbook = null;
        try {
            if ("xls".equals(type)) {
                workbook = new HSSFWorkbook(new FileInputStream(file));
            }
            if ("xlsx".equals(type)) {
                workbook = new XSSFWorkbook(new FileInputStream(file));
            }
        } catch (Exception e) {
            e.printStackTrace();
            log.info("获取Excel表的WorkBook失败");
        }
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if(!workbook.isSheetHidden(i)) {
                sheetList.add(workbook.getSheetAt(i));
            }
        }
        return sheetList.iterator();
    }

    /**
     * 表头判断
     *
     * @param cell
     * @return
     */
    public static boolean isHead(Sheet sheet, Cell cell, ConfigBean configBean, String type) {
        Cell nextCell = getHeadNextRangeCell(sheet, cell, configBean);
        // 判断表头是否需要精确匹配
        try {
            Header header = configBean.getHeader();
            boolean headAccurate = header.getAccurate();
            String headColor = header.getBgColor();
            String[] headPattern = header.getPattern();
            String[] headStyle = header.getStyle();
            boolean colorMatchResult = MatchUtil.colorMatch(getColorByCell(cell, type), headColor);
            boolean patternMatchResult = MatchUtil.patternMatch(cell.getStringCellValue(), headPattern);
            boolean styleMatchResult = MatchUtil.styleMatch(cell, nextCell, headStyle);
            return headAccurate ? colorMatchResult && patternMatchResult && styleMatchResult
                    : colorMatchResult || patternMatchResult || styleMatchResult;
        } catch (Exception ignored) {
            return false;
        }
    }

    /**
     * 表尾判断
     * @param cell
     * @return
     */
    public static boolean isTail(Sheet sheet, Cell cell, ConfigBean configBean, String type) {
        Cell nextCell = getTailNextRangeCell(sheet, cell, configBean);
        // 判断表尾是否需要精确匹配
        try {
            Tail tail = configBean.getTail();
            boolean tailAccurate = tail.getAccurate();
            String tailColor = tail.getBgColor();
            String[] tailPattern = tail.getPattern();
            String[] tailStyle = tail.getStyle();
            String fontColor = tail.getFontColor();
            boolean colorMatchResult = MatchUtil.colorMatch(getColorByCell(cell, type), tailColor);
            boolean patternMatchResult = MatchUtil.patternMatch(cell.getStringCellValue(), tailPattern);
            boolean styleMatchResult = MatchUtil.styleMatch(cell, nextCell, tailStyle);
            boolean fontColorMatchResult = MatchUtil.fontColorMatch(getColorByFont(cell), fontColor);
            return tailAccurate ? colorMatchResult && patternMatchResult && styleMatchResult && fontColorMatchResult
                    : colorMatchResult || patternMatchResult || styleMatchResult || fontColorMatchResult;
        } catch (Exception ignored) {
            return false;
        }
    }

    /** 标题判断
     * @param cell
     * @return
     */
    public static boolean isTitle(Sheet sheet, Cell cell, ConfigBean configBean, String type) {

        Cell nextCell = getTailNextRangeCell(sheet, cell, configBean);
        try {
            TitleFeature titleFeature = configBean.getContent().getTitleFeature();
            String titleColor = titleFeature.getColor();
            String[] titlePattern = titleFeature.getPattern();
            String cellColor = getColorByCell(cell, type);
            String[] style = titleFeature.getStyle();
            boolean colorMatch = MatchUtil.colorMatch(cellColor, titleColor);
            boolean patternMatch = MatchUtil.patternMatch(cell.getStringCellValue(), titlePattern);
            boolean styleMatch = MatchUtil.styleMatch(cell, nextCell, style);
            return colorMatch || patternMatch || styleMatch;
        } catch (Exception ignored) {
            return false;
        }
    }

    /**
     * 获取表头当前单元格下一行对应的单元格
     * @param sheet
     * @param cell
     * @return
     */
    private static Cell getHeadNextRangeCell(Sheet sheet, Cell cell, ConfigBean configBean) {
        try {
            Integer range = configBean.getHeader().getRowRange();
            Row nextRow = sheet.getRow(cell.getRowIndex() + range);
            Cell nextCell = nextRow.getCell(cell.getColumnIndex());
            return nextCell;
        } catch (Exception ignored) {
            return null;
        }
    }

    /**
     * 获取表尾当前单元格下一行对应的单元格
     * @param sheet
     * @param cell
     * @return
     */
    private static Cell getTailNextRangeCell(Sheet sheet, Cell cell, ConfigBean configBean) {
        try {
            Integer range = configBean.getTail().getRowRange();
            Row nextRow = sheet.getRow(cell.getRowIndex() - range);
            Cell nextCell = nextRow.getCell(cell.getColumnIndex());
            return nextCell;
        } catch (Exception ignored) {
            return null;
        }
    }

    /**
     * 获取颜色
     *
     * @param cell
     * @return
     */
    public static String getColorByCell(Cell cell, String type) {
        StringBuilder sb = new StringBuilder();
        CellStyle style = cell.getCellStyle();

        if ("xls".equals(type)) {
            short color = style.getFillForegroundColor();
            HSSFWorkbook hb = new HSSFWorkbook();
            HSSFColor hc = hb.getCustomPalette().getColor(color);
            short[] s = hc.getTriplet();
            if (s != null) {
                sb.append(s[0]).append(",").append(s[1]).append(",").append(s[2]);
            }
        } else {


            XSSFColor color = (XSSFColor) style.getFillForegroundColorColor();
            if (color != null) {
                if (color.isRGB()) {
                    byte[] bytes = color.getRGB();
                    if (bytes != null && bytes.length == 3) {
                        for (int i = 0; i < bytes.length; i++) {
                            byte b = bytes[i];
                            int temp;
                            if (b < 0) {
                                temp = 256 + (int) b;
                            } else {
                                temp = b;
                            }
                            sb.append(temp);
                            if (i != bytes.length - 1) {
                                sb.append(",");
                            }
                        }
                    }
                }
            }
        }
        return sb.toString();

    }

    /**
     * 获取字体颜色
     * @param cell
     * @return
     */
    public static String getColorByFont(Cell cell) {
        return String.valueOf(cell.getCellStyle().getFontIndex());
    }

    /**
     * 获取转换后的单元格值
     * @param cell
     * @return
     */
    public static String getCellConvertValue(Cell cell) {
        String cellValue= "";
        CellType cellType;
        try {
            cellType = cell.getCellTypeEnum();
        } catch (Exception e) {
            cellType = CellType.STRING;
        }
        switch(cellType) {
            case STRING :
                cellValue = cell.getStringCellValue().trim();
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    cellValue = DateUtils.format(cell.getDateCellValue(), "yyyy-MM-dd");
                } else {
                    cellValue = new DecimalFormat("#.######").format(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:

                try {
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        cellValue = DateUtils.format(cell.getDateCellValue(), "yyyy-MM-dd");
                    } else {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                } catch (IllegalStateException e) {
                    cell.setCellType(CellType.STRING);
                    cellValue = String.valueOf(cell.getRichStringCellValue());
                }
                break;
            default:
                cellValue = "";
                break;
        }
        return cellValue;

    }

    /**
     * 判断指定的单元格是否是合并单元格
     * @param sheet
     * @param row 行下标
     * @param column 列下标
     * @return
     */
    public static boolean isMergedRegion(Sheet sheet,int row ,int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 初始化范围
     * @param range
     */
    public static void initRange(List<Integer> range) {
        if (range.size() == 0) {
            range.add(0);
            range.add(0);
        }
        else {
            range.set(0, 0);
            range.set(1, 0);
        }
    }

    /**
     * 解析忽略的列
     * @param row
     * @return
     */
    public static Set<Integer> getIgnoredColumn(Set<Integer> ignoredColumn, Row row, ConfigBean configBean) {
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            String cellValue = getCellConvertValue(cell);
            String[] patterns = configBean.getContent().getIgnoreColumn();
            for (String pattern : patterns) {
                Pattern p = Pattern.compile(pattern);
                Matcher matcher = p.matcher(cellValue);
                if (matcher.find()) {
                    ignoredColumn.add(cell.getColumnIndex());
                }
            }
        }
        return ignoredColumn;
    }

    /**
     * 解析标题
     * @param portOfCallRange
     * @param specifiedAttribute
     */
    public static void parseTitle(List<Integer> portOfCallRange, VslVoyAttribute specifiedAttribute, Row row) {
        int begin = portOfCallRange.get(0);
        int end = portOfCallRange.get(1);
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            if (getCellConvertValue(cell).contains(specifiedAttribute.getBegin())) {
                begin = i + 1;
            }
            else if (getCellConvertValue(cell).contains(specifiedAttribute.getEnd())) {
                end = i - 1;
            }
        }
        portOfCallRange.set(0, begin);
        portOfCallRange.set(1, end);
    }

    /**
     * 解析标题中的合并单元格
     * @param sheet
     * @param temp
     * @param cellNum
     * @param mergedTitle
     */
    public static void parseMergedInTitle(Sheet sheet, Cell temp, Integer cellNum, Map<Integer,String> mergedTitle) {
        // 判断是否为合并单元格
        boolean mergedRegionFlag = ExcelUtil.isMergedRegion(sheet, temp.getRowIndex(), temp.getColumnIndex());
        if (mergedRegionFlag) {
            mergedTitle.put(cellNum - 1, ExcelUtil.getCellConvertValue(temp));
        }
    }

    /**
     * 获取标题中的挂靠港名称
     * @param portOfCallList
     * @param portOfCallRange
     * @param row
     */
    public static void getPortOfCallNameInTitle(List<PortOfCall> portOfCallList, List<Integer> portOfCallRange, Row row) {
        for (int i = portOfCallRange.get(0); i <= portOfCallRange.get(1); i++) {
            PortOfCall portOfCall = new PortOfCall();
            Cell cell = row.getCell(i);
            String cellValue = getCellConvertValue(cell);
            portOfCall.setPortOfCallName(cellValue);
            portOfCallList.add(portOfCall);
        }
    }

    /**
     * 获取在表头里的航线名称
     * @param header
     * @return
     */
    public static String getLineNameInHeader(String header) {
        String res = "";

        Map<String, String> lineMapper = LineMapper.getLineMapper();
        LineMapper.init();
        Set<String> lineNames = lineMapper.keySet();
        char black = ' ';
        char englishBracket = '(';
        char chineseBracket = '（';
        char slash = '/';
        // 针对各种格式的表头，采用不同的解析方法
        if (StringUtils.contains(header, black) && (StringUtils.contains(header, englishBracket) || StringUtils.contains(header, chineseBracket) )) {
            List<Character> splitFlag = new ArrayList<>();
            splitFlag.add(' ');
            splitFlag.add('(');
            splitFlag.add(')');
            splitFlag.add('（');
            splitFlag.add('）');
            res = splitHeader(header, splitFlag, lineNames);
        }
        else if (!StringUtils.contains(header, black) && (StringUtils.contains(header, englishBracket) || StringUtils.contains(header, chineseBracket) )) {
            List<Character> splitFlag = new ArrayList<>();
            splitFlag.add('(');
            splitFlag.add(')');
            splitFlag.add('（');
            splitFlag.add('）');
            res = splitHeader(header, splitFlag, lineNames);
        }
        else if (StringUtils.contains(header, black) && !(StringUtils.contains(header, englishBracket) || StringUtils.contains(header, chineseBracket) )) {
            List<Character> splitFlag = new ArrayList<>();
            splitFlag.add(' ');
            res = splitHeader(header, splitFlag, lineNames);
        }
        else if (!StringUtils.contains(header, black) && !(StringUtils.contains(header, englishBracket) || StringUtils.contains(header, chineseBracket) ) && !StringUtils.contains(header, slash)) {
            res = header;
        }
        else if (StringUtils.contains(header, slash)) {
            List<Character> splitFlag = new ArrayList<>();
            splitFlag.add('/');
            res = splitHeader(header, splitFlag, lineNames);
        }
        return res;

    }

    /**
     * 分割表头
     * @param header
     * @param splitFlag
     * @param lineNames
     * @return
     */
    private static String splitHeader(String header, List<Character> splitFlag, Set<String> lineNames) {
        String line = "";
        List<String> res = new ArrayList<>();
        List<String> convert = new ArrayList<>();
        StringBuilder temp = new StringBuilder();
        for (int i = 0; i < header.length(); i++) {

            if (!splitFlag.contains(header.charAt(i))) {
                temp.append(header.charAt(i));
            }
            else {
                res.add(temp.toString());
                temp = new StringBuilder();
            }
            if (i == header.length() - 1) {
                res.add(temp.toString());
                temp = new StringBuilder();
            }
        }
        for (String str : res) {
            str = removeBrackets(str);
            convert.add(str);
        }
        for (String str : convert) {
            str = str.trim();
            if (lineNames.contains(str)) {
                line = str;
                break;
            }
        }
        return line;
    }

    /**
     * 去除括号
     * @param str
     * @return
     */
    private static String removeBrackets(String str) {
        StringBuilder res = new StringBuilder();
        for (int i = 0; i < str.length(); i++) {
            if ('(' == str.charAt(i) || ')' == str.charAt(i) || '（' == str.charAt(i) || '）' == str.charAt(i)) {
                continue;
            }
            res.append(str.charAt(i));
        }
        return res.toString();
    }

}
