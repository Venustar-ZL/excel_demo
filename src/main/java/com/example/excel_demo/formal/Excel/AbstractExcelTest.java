package com.example.excel_demo.formal.Excel;

import com.alibaba.fastjson.JSONObject;
import com.example.excel_demo.formal.bean.Line;
import com.example.excel_demo.formal.bean.PortOfCall;
import com.example.excel_demo.formal.bean.VslVoy;
import com.example.excel_demo.formal.bean.VslVoys;
import com.example.excel_demo.formal.config.ConfigBean;
import com.example.excel_demo.formal.config.VslVoyAttribute;
import com.example.excel_demo.formal.mapper.LineMapper;
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
public class AbstractExcelTest {

    private static ConfigBean configBean;

    public static void main(String[] args) {
        File file = new File("C:\\Users\\hujingyi\\Desktop\\2月船期.xlsx");
        parseNormalExcel(file, "application-DNY");
    }

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
            if (!sheet.getSheetName().contains("东南亚")) {
                continue;
            }
            parseSheet(sheet, type);
        }

    }

    private static List<Line> parseSheet(Sheet sheet, String type) {

        Set<Integer> errorRow = new HashSet<>();
        int errorCount = 0;

        List<Line> lineList = new ArrayList<>();
        Line line = null;
        VslVoys vslVoys = null;
        List<VslVoy> vslVoyList = new ArrayList<>();

        List<VslVoyAttribute> attributeList = new ArrayList<>();
        VslVoyAttribute specialAttribute = new VslVoyAttribute();
        Set<Integer> ignoredColumn = new HashSet<>();
        List<Integer> portOfCallRange = new ArrayList<>();
        List<PortOfCall> portOfCallList = new ArrayList<>();
        Map<Integer, String> mergedCell = new HashMap<>(16);
        StringBuilder remark = new StringBuilder();
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

                        vslVoyList.clear();
                        line = new Line();
                        vslVoys = new VslVoys();
                        remark = new StringBuilder();
                        dealHead(cell, mergedCell, portOfCallList, portOfCallRange, ignoredColumn);
                        convertConfigBean(line, ExcelUtil.getCellConvertValue(cell));
                        attributeList = YamlUtil.parseUniversal(configBean);
                        specialAttribute = YamlUtil.parseSpecial(configBean);
                        contentFlag = true;
                        continue;
                    }

                    if (ExcelUtil.isTail(sheet, cell, configBean, type)) {
                        dealTail(cell, mergedCell, portOfCallList, portOfCallRange, ignoredColumn);
                        remark.append(ExcelUtil.getCellConvertValue(cell));
                        addRemarks(vslVoys, remark.toString());
                        addLine(vslVoyList, vslVoys, lineList, line);
                        contentFlag = false;
                        continue;
                    }


                    if (contentFlag) {

                        if (StringUtils.isBlank(ExcelUtil.getCellConvertValue(cell))) {
                            continue;
                        }

                        if (ExcelUtil.isTitle(sheet, cell, configBean, type)) {
                            dealTitle(row, ignoredColumn, portOfCallList, portOfCallRange, specialAttribute);
                            continue;
                        }

                        dealContent(vslVoy, attributeList, portOfCallList, portOfCallRange, ignoredColumn, mergedCell, sheet, row, specialAttribute, cellNum);
                        if (StringUtils.isNotBlank(vslVoy.getVesselName())) {
                            vslVoyList.add(vslVoy);
                            System.out.println(JSONObject.toJSONString(vslVoy));
                        }
                    }

                }
            } catch (Exception e) {
                e.printStackTrace();
                log.error("第[{}]行解析失败", rowNum + 1);
                errorCount++;
                errorRow.add(rowNum + 1);
            }
        }

        System.out.println("------------------------------");
        System.out.println("航线条数：" + lineList.size());
        System.out.println(JSONObject.toJSONString(lineList));
        System.out.println("失败行数 : " + errorCount);
        System.out.println("具体行数为 : " + Arrays.toString(errorRow.toArray()));
        System.out.println("------------------------------");

        for (Line line1 : lineList) {
            System.out.println(JSONObject.toJSONString(line1));
        }

        return lineList;
    }


    public static void dealHead(Cell cell, Map<Integer, String> mergedCell, List<PortOfCall> portOfCallList, List<Integer> portOfCallRange, Set<Integer> ignoredColumn) {

        log.info("=====> 表头 <=====");
        log.info(cell.getStringCellValue());
        ignoredColumn.clear();
        mergedCell.clear();
        portOfCallList.clear();
        ExcelUtil.initRange(portOfCallRange);

    }

    public static void dealTail(Cell cell, Map<Integer, String> mergedCell, List<PortOfCall> portOfCallList, List<Integer> portOfCallRange, Set<Integer> ignoredColumn) {

        log.info("=====> 表尾 <=====");
        log.info(cell.getStringCellValue());
        ignoredColumn.clear();
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
                if (rangeFlag) {
                    portOfCallNo++;
                    InvokeUtil.dynamicListAdd(vslVoy, cellValue, portOfCallNo, portOfCallList);
                    k--;
                }
            }

        }
        return cellNum;

    }

    private static void convertConfigBean(Line line, String header) {
        if (StringUtils.isBlank(header)) {
            return;
        }
        String lineName = ExcelUtil.getLineNameInHeader(header);
        line.setLineName(lineName);
        log.info("航线名称：[{}]", lineName);
        String yamlName = LineMapper.getLineMapper().get(lineName);
        log.info("加载的配置文件名称：[{}]", yamlName);
        configBean = YamlUtil.getConfigBean(yamlName);
    }

    /**
     * 添加备注信息
     * @param vslVoys
     * @param cellValue
     */
    private static void addRemarks(VslVoys vslVoys, String cellValue) {
        if (StringUtils.isBlank(cellValue)) {
            return;
        }
        vslVoys.setRemarks(cellValue + "\n");
    }

    private static void addLine(List<VslVoy> vslVoyList, VslVoys vslVoys, List<Line> lineList, Line line) {
        if (vslVoyList.size() != 0 && vslVoys.getVslVoyList() == null) {
            vslVoys.setVslVoyList(vslVoyList);
            line.setVslVoys(vslVoys);
            lineList.add(line);
        }
    }

}
