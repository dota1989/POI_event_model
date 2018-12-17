package com.zb.poi_event_model.normal;

import com.zb.poi_event_model.event.ExcelReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by zhangbin on 2018/11/8.
 */
public abstract class ExcelUtil {
    private static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    private int size = 1000;
   /**
     * 解析excel文件
     *
     * @param fileName
     * @return
     * @throws IOException
     */
    public int parseExcel(String fileName) throws IOException {
        logger.info("解析Excel文件：{}", fileName);

        FileInputStream fileIn = new FileInputStream(fileName);

        //根据指定的文件输入流导入Excel从而产生Workbook对象
        Workbook wb;
        if (fileName.endsWith("xlsx")) {
            wb = new XSSFWorkbook(fileIn);
        } else {
            wb = new HSSFWorkbook(fileIn);
        }

        Sheet sheet = wb.getSheetAt(0);

        //首行
        Row firstRow = sheet.getRow(0);
        List<String> keyList = new ArrayList<>();
        for (Cell cell : firstRow) {
            keyList.add(getCellValue(cell));
        }

        //临时集合
        List<Map<String, String>> mapList = new ArrayList<>();
        List<String> datalist = new ArrayList<>();

        //内容
        int total = sheet.getLastRowNum();
        int scanCount = total/size + 1;
        for (int k = 0; k < scanCount; k++) {
            for (int i = k*size; i < (k+1)*size; i++) {
                Row row = sheet.getRow(i);
                if(row == null) continue;
                int j = 0;
                Map<String, String> map = new HashMap<>();
                for (int l = 0; l < keyList.size(); l++) {
                    Cell cell = row.getCell(l, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (cell == null || cell.toString() == null) {
                        j++;
                        continue;
                    }
                    map.put(keyList.get(j), getCellValue(cell));
                    j++;
                }
                mapList.add(map);
            }

            //插入数据
            //在这里进行数据处理，数据量不大的话也可以统一处理
            processData(mapList);

            mapList.clear();
            datalist.clear();
        }

        //处理剩余数据
        processData(mapList);
        return total;
    }

    abstract void processData(List<Map<String, String>> mapList);

    /**
     * 获得单元格的值
     *
     * @param cell 单元格
     * @return
     */
    private String getCellValue(Cell cell) {
        CellType cellType = cell.getCellTypeEnum();
        String cellValue = "";
        switch (cellType) {
            case NUMERIC:
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                cellValue = cell.getCellFormula();
                break;
            case BLANK:
                cellValue = "";
                break;
            case ERROR:
                cellValue = String.valueOf(cell.getErrorCellValue());
                break;
            default:
                cellValue = "";
        }
        return cellValue;
    }

    /**
     * 测试方法
     */
    public static void main(String[] args) throws Exception {
/////////////////////////////////////////////////////////////////////////////////////////////////////////
        try {
            ExcelUtil reader = new ExcelUtil() {

                @Override
                void processData(List<Map<String, String>> mapList) {
                    System.out.println(mapList);
                }
            };
            reader.parseExcel("src/main/resources/22.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }
/////////////////////////////////////////////////////////////////////////////////////////////////////////
    }
}
