package com.zb.poi_event_model.normal;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.util.*;

/**
 * Created by zhangbin on 2018/11/28.
 */
@Component
public class ExportExcelUtil {
    private static final Logger logger = LoggerFactory.getLogger(ExportExcelUtil.class);

    public Workbook exportExcelFile(List<Map<String, Object>> resultDataList) {
        try {
            logger.info("生成Excel文件");
//            Workbook workbook = new XSSFWorkbook();
            SXSSFWorkbook workbook = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
            Sheet sheet = workbook.createSheet("挂牌案例数据");

            // 设置单元格格式为文本格式
            DataFormat dataFormat = workbook.createDataFormat();
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setDataFormat(dataFormat.getFormat("@"));

            Map<String, Integer> columnIndexMap = new HashMap<>();
            int columnIndex = 0;
            int rowIndex = 0;
            Row headRow = sheet.createRow(rowIndex++);

            // 表头
//            List<StandardColumn> standardColumnList = standardColumnService.findAll();
//            for (StandardColumn standardColumn: standardColumnList) {
//                Cell cell = headRow.createCell(columnIndex);
//                cell.setCellValue(standardColumn.getCndescribe());
//                columnIndexMap.put(standardColumn.getCncolumn(), columnIndex);
//                if("cleanstatus".equals(standardColumn.getEncolumn())
//                        || "cleanmsg".equals(standardColumn.getEncolumn())
//                        || "storestatus".equals(standardColumn.getEncolumn())
//                        || "storemsg".equals(standardColumn.getEncolumn())){
//                    columnIndexMap.put(standardColumn.getEncolumn(), columnIndex);
//                }
//                columnIndex++;
//            }

            // 数据
            for (Map<String, Object> resultMap: resultDataList) {
                Row contentRow = sheet.createRow(rowIndex++);
                for (Map.Entry<String, Object> entry: resultMap.entrySet()) {
                    Integer index = columnIndexMap.get(entry.getKey());
                    if (index == null){
                        continue;
                    }
                    Cell cell = contentRow.createCell(index);
                    cell.setCellStyle(cellStyle);
                    setCellValue(cell, entry.getValue());
                }
            }
            return workbook;
        } catch (Exception e){
            logger.error("exportExcelFile: ", e);
            return null;
        }
    }

    /**
     * 设置单元格内容
     *
     * @param cell
     * @param value
     */
    private void setCellValue(Cell cell, Object value) {
        if(value == null){
            return;
        }
        if(value instanceof Double){
            cell.setCellValue((double)value);
            cell.setCellValue(new BigDecimal(value.toString()).toPlainString());
        }
        if(value instanceof Integer){
            cell.setCellValue(new BigDecimal(value.toString()).toPlainString());
        }
        if(value instanceof Long){
            cell.setCellValue(new BigDecimal(value.toString()).toPlainString());
        }
        if(value instanceof String){
            cell.setCellValue((String)value);
        }
        if(value instanceof Boolean){
            cell.setCellValue((Boolean)value);
        }
        if(value instanceof Date){
            cell.setCellValue((Date)value);
        }
        if(value instanceof Calendar){
            cell.setCellValue((Calendar)value);
        }
        if(value instanceof RichTextString){
            cell.setCellValue((RichTextString)value);
        }
    }
}
