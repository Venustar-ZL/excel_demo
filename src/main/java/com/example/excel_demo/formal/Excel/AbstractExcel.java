package com.example.excel_demo.formal.Excel;

import com.alibaba.fastjson.JSONObject;
import com.example.excel_demo.formal.bean.PortOfCall;
import com.example.excel_demo.formal.bean.VslVoy;
import com.example.excel_demo.formal.config.ConfigBean;
import com.example.excel_demo.formal.config.VslVoyAttribute;
import com.example.excel_demo.formal.utils.ExcelUtil;
import com.example.excel_demo.formal.utils.InvokeUtil;
import com.example.excel_demo.formal.utils.YamlUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.util.*;

/**
 * @Classname AbstractExcel
 * @Description TODO
 * @Date 2020/12/9 16:39
 * @Author by ZhangLei
 */
@Slf4j
public class AbstractExcel {

    public static void main(String[] args) {
        File file = new File("C:\\Users\\hujingyi\\Desktop\\2月船期.xlsx");
        parseNormalExcel(file, "application-YT");
    }

    private static ConfigBean configBean;

    public static void parseNormalExcel(File file, String yamlName) {
        // 获取配置
        configBean = YamlUtil.getConfigBean(yamlName);

        // 解析Excel文件
        // 获取文件类型
        String type = ExcelUtil.getExcelType(file);
        // 获取所有不隐藏的sheet
        Iterator<Sheet> sheets = ExcelUtil.getAllNotHiddenSheet(file, type);
        if (sheets == null) {
            log.info("Excel中不含有sheet工作表");
        }
        // 读取sheet
        while (sheets.hasNext()) {
            Sheet sheet = sheets.next();
            if (!sheet.getSheetName().contains("亚太")) {
                continue;
            }
            parseSheet(sheet, type);
        }

    }

    private static List<VslVoy> parseSheet(Sheet sheet, String type) {

        List<VslVoy> list = new ArrayList<>();
        List<VslVoyAttribute> attributeList = YamlUtil.parseUniversal(configBean);
        VslVoyAttribute specialAttribute = YamlUtil.parseSpecial(configBean);
        Set<Integer> ignoredColumn = new HashSet<>();
        List<Integer> portOfCallRange = new ArrayList<>();
        List<PortOfCall> portOfCallList = new ArrayList<>();
        Map<Integer, String> mergedCell = new HashMap<>(16);
        Map<Integer, String> mergerTitle = new HashMap<>(16);
        ExcelUtil.initRange(portOfCallRange);
        boolean contentFlag = false;
        for (int rowNum = sheet.getFirstRowNum(); rowNum < sheet.getLastRowNum(); rowNum++) {
            VslVoy vslVoy = new VslVoy();
            List<PortOfCall> portOfCalls = new ArrayList<>();
            vslVoy.setPortOfCalls(portOfCalls);
            Row row = sheet.getRow(rowNum);

            try {
                if (row != null) {
                    int cellNum = row.getFirstCellNum();
                    Cell cell = row.getCell(cellNum);

                    if (ExcelUtil.isHead(sheet, cell, configBean, type)) {
                        dealHead(cell, mergedCell, portOfCallList, portOfCallRange);
                        contentFlag = true;
                        continue;
                    }

                    if (ExcelUtil.isTail(sheet, cell, configBean, type)) {
                        dealTail(cell, mergedCell, portOfCallList, portOfCallRange);
                        contentFlag = false;
                        continue;
                    }


                    if (contentFlag) {

                        if (StringUtils.isBlank(ExcelUtil.getCellConvertValue(cell))) {
                            continue;
                        }

//                        log.info("行数[{}]", cell.getRowIndex() + " 单元格值" + cell.getStringCellValue());
                        if (ExcelUtil.isTitle(sheet, cell, configBean, type)) {
                            dealTitle(row, ignoredColumn, portOfCallList, portOfCallRange, specialAttribute);
                            continue;
                        }

                        dealContent(vslVoy, attributeList, portOfCallList, portOfCallRange, ignoredColumn, mergedCell, sheet, row, specialAttribute, cellNum);
                    }

                }
            } catch (Exception e) {
                log.error("第[{}]行解析失败", rowNum + 1);
            }
            if (StringUtils.isNotBlank(vslVoy.getVesselName())) {
                System.out.println(JSONObject.toJSONString(vslVoy));
            }
        }

        return list;
    }


    public static void dealHead(Cell cell, Map<Integer, String> mergedCell, List<PortOfCall> portOfCallList, List<Integer> portOfCallRange) {

        log.info("=====> 表头 <=====");
        log.info(cell.getStringCellValue());
        mergedCell.clear();
        portOfCallList.clear();
        ExcelUtil.initRange(portOfCallRange);

    }

    public static void dealTail(Cell cell, Map<Integer, String> mergedCell, List<PortOfCall> portOfCallList, List<Integer> portOfCallRange) {

        log.info("=====> 表尾 <=====");
        log.info(cell.getStringCellValue());
        mergedCell.clear();
        portOfCallList.clear();
        ExcelUtil.initRange(portOfCallRange);

    }

    public static void dealTitle(Row row, Set<Integer> ignoredColumn, List<PortOfCall> portOfCallList, List<Integer> portOfCallRange, VslVoyAttribute specialAttribute) {
        ExcelUtil.getIgnoredColumn(ignoredColumn, row, configBean);
        ExcelUtil.parseTitle(portOfCallRange, specialAttribute, row);
        ExcelUtil.getPortOfCallNameInTitle(portOfCallList, portOfCallRange, row);
    }

    public static Integer dealContent(VslVoy vslVoy, List<VslVoyAttribute> attributeList, List<PortOfCall> portOfCallList, List<Integer> portOfCallRange, Set<Integer> ignoredColumn, Map<Integer, String> mergedCell, Sheet sheet, Row row, VslVoyAttribute specialAttribute, Integer cellNum) {

        boolean specialBeginFlag = false;
        int portOfCallNo = 0;

        for (int k = 0; k < attributeList.size(); k++) {

            // 判断此列是否需要忽略
            if (ignoredColumn.contains(cellNum)) {
                cellNum++;
                k--;
                continue;
            }

            String cellValue = "";
            if (mergedCell.containsKey(cellNum)) {
                cellValue = mergedCell.get(cellNum);
                cellNum++;
            } else {

                Cell temp = row.getCell(cellNum++);
                if (temp == null) {
                    continue;
                }

                // 判断是否为合并单元格
                boolean mergedRegionFlag = ExcelUtil.isMergedRegion(sheet, temp.getRowIndex(), temp.getColumnIndex());
                if (mergedRegionFlag) {
                    mergedCell.put(cellNum - 1, ExcelUtil.getCellConvertValue(temp));
                }

                cellValue = ExcelUtil.getCellConvertValue(temp);
            }

            VslVoyAttribute universalAttribute = attributeList.get(k);

            // 判断是否进入特殊属性范围,则需进行特殊属性的处理，处理完成之后，不影响通用属性的处理顺序
            if (universalAttribute.getName().equals(specialAttribute.getBegin())) {
                specialBeginFlag = true;
            }

            boolean rangeFlag = ((cellNum >= portOfCallRange.get(0) + 1 && cellNum <= portOfCallRange.get(1) + 1));
            if (universalAttribute.getLength() > 1 && !rangeFlag) {
                for (int start = cellNum; start < cellNum + universalAttribute.getLength(); start++) {
                    Cell c = row.getCell(start);
                    if (c == null) {
                        continue;
                    }
                    String value = ExcelUtil.getCellConvertValue(c);
                    cellValue = cellValue + " " + value;
                }
                cellNum++;
            }

            InvokeUtil.dynamicSet(vslVoy, universalAttribute.getName(), cellValue);
            cellNum = cellNum + universalAttribute.getLength() - 1;

            if ("null".equals(specialAttribute.getEnd())) {
                if (specialBeginFlag) {
                    portOfCallNo++;
                    InvokeUtil.dynamicListAdd(vslVoy, cellValue, portOfCallNo, portOfCallList);
                    k--;
                    cellNum++;
                }
            } else {
                if (cellNum >= portOfCallRange.get(0) + 1 && cellNum <= portOfCallRange.get(1) + 1) {
                    portOfCallNo++;
                    InvokeUtil.dynamicListAdd(vslVoy, cellValue, portOfCallNo, portOfCallList);
                    k--;
                }
            }

        }
        return cellNum;

    }

    /**
     * 默认由配置文件导入，只对类型进行校验，复杂逻辑校验可由子类覆盖实现
     * @return
     */
    private static boolean check() {}

}
