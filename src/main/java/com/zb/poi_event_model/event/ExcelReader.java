package com.zb.poi_event_model.event;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Excel超大数据读取，抽象Excel2007读取器，excel2007的底层数据结构是xml文件，采用SAX的事件驱动的方法解析
 * xml，需要继承DefaultHandler，在遇到文件内容时，事件会触发，这种做法可以大大降低 内存的耗费，特别使用于大数据量的文件。
 * @version 2014-9-2
 */
public abstract class ExcelReader extends DefaultHandler {

    // 共享字符串表
    private SharedStringsTable sst;

    // 上一次的内容
    private String lastContents;
    private boolean nextIsString;

    private int sheetIndex = -1;

    // 当前行
    private int curRow = 0;
    // 日期标志
    private boolean dateFlag;
    // 数字标志
    private boolean numberFlag;

    private boolean isTElement;
    // 当前列
    private String columnStr = "";
    // 每一行的数据
    private Map<String, String> dataMap = new HashMap<>();


    /**
     * 单元格
     */
    private StylesTable stylesTable;

    /**
     * 单元格数据类型，默认为字符串类型
     */
    private CellDataType nextDataType = CellDataType.SSTINDEX;

    private final DataFormatter formatter = new DataFormatter();

    private short formatIndex;

    private String formatString;

    /**
     * 遍历工作簿中所有的电子表格
     * @param filename
     * @throws Exception
     */
    public void process(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader(pkg);
        stylesTable = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        Iterator<InputStream> sheets = r.getSheetsData();
        while (sheets.hasNext()) {
            curRow = 0;
            sheetIndex++;
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
        }
    }

    /**
     * 只遍历一个电子表格，其中sheetId为要遍历的sheet索引，从1开始，1-3
     * @param filename
     * @param sheetId
     * @throws Exception
     */
    public void process(String filename, int sheetId) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader(pkg);
        stylesTable = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        // 根据 rId# 或 rSheet# 查找sheet
        InputStream sheet2 = r.getSheet("rId" + sheetId);
        sheetIndex++;
        InputSource sheetSource = new InputSource(sheet2);
        parser.parse(sheetSource);
        sheet2.close();
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst)
            throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        this.sst = sst;
        parser.setContentHandler(this);
        return parser;
    }

    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {

//      System.out.println("startElement: " + localName + ", " + name + ", " + attributes);
        // c => 单元格
        if ("c".equals(name)) {
            String indexStr = attributes.getValue("r");
            columnStr = parseIndexStr(indexStr);
            // 如果下一个元素是 SST 的索引，则将nextIsString标记为true
            String cellType = attributes.getValue("t");
            if ("s".equals(cellType)) {
                nextIsString = true;
            } else {
                nextIsString = false;
            }
            // 设定单元格类型
            this.setNextDataType(attributes);
        }
        // 当元素为t时
        if ("t".equals(name)) {
            isTElement = true;
        } else {
            isTElement = false;
        }

        // 置空
        lastContents = "";
    }

    public void endElement(String uri, String localName, String name)
            throws SAXException {
//      System.out.println("endElement: " + localName + ", " + name);

        // 根据SST的索引值的到单元格的真正要存储的字符串
        // 这时characters()方法可能会被调用多次
        if (nextIsString) {
            try {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx))
                        .toString();
            } catch (Exception e) {

            }
        }
        // t元素也包含字符串
        if (isTElement) {
            String value = lastContents.trim();
            dataMap.put(columnStr, value);
            isTElement = false;
            // v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引
            // 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
        } else if ("v".equals(name)) {
            String value = this.getDataValue(lastContents.trim(), "");
            dataMap.put(columnStr, value);
        } else {
            // 如果标签名称为 row ，这说明已到行尾，调用 optRows() 方法
            if (name.equals("row")) {
                getRows(sheetIndex + 1, curRow, dataMap);
                dataMap.clear();
                curRow++;
            }
        }
    }

    /**
     * 单元格中的数据可能的数据类型
     */
    enum CellDataType {
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
    }

    /**
     * 处理数据类型
     *
     * @param attributes
     */
    public void setNextDataType(Attributes attributes) {
        nextDataType = CellDataType.NUMBER;
        formatIndex = -1;
        formatString = null;
        String cellType = attributes.getValue("t");
        String cellStyleStr = attributes.getValue("s");

        if ("b".equals(cellType)) {
            nextDataType = CellDataType.BOOL;
        } else if ("e".equals(cellType)) {
            nextDataType = CellDataType.ERROR;
        } else if ("inlineStr".equals(cellType)) {
            nextDataType = CellDataType.INLINESTR;
        } else if ("s".equals(cellType)) {
            nextDataType = CellDataType.SSTINDEX;
        } else if ("str".equals(cellType)) {
            nextDataType = CellDataType.FORMULA;
        }

        if (cellStyleStr != null) {
            int styleIndex = Integer.parseInt(cellStyleStr);
            XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
            formatIndex = style.getDataFormat();
            formatString = style.getDataFormatString();

            if ("m/d/yy" == formatString) {
                nextDataType = CellDataType.DATE;
                formatString = "yyyy-MM-dd";
//                formatString = "yyyy-MM-dd hh:mm:ss.SSS";
            }

            if (formatString == null) {
                nextDataType = CellDataType.NULL;
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }
        }
    }

    /**
     * 对解析出来的数据进行类型处理
     *
     * @param value
     *            单元格的值（这时候是一串数字）
     * @param thisStr
     *            一个空字符串
     * @return
     */
    @SuppressWarnings("deprecation")
    public String getDataValue(String value, String thisStr) {
        switch (nextDataType) {
            // 这几个的顺序不能随便交换，交换了很可能会导致数据错误
            case BOOL:
                char first = value.charAt(0);
                thisStr = first == '0' ? "FALSE" : "TRUE";
                break;
            case ERROR:
                thisStr = "\"ERROR:" + value.toString() + '"';
                break;
            case FORMULA:
                thisStr = '"' + value.toString() + '"';
                break;
            case INLINESTR:
                XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());

                thisStr = rtsi.toString();
                rtsi = null;
                break;
            case SSTINDEX:
                String sstIndex = value.toString();
                try {
                    int idx = Integer.parseInt(sstIndex);
                    XSSFRichTextString rtss = new XSSFRichTextString(sst.getEntryAt(idx));
                    thisStr = rtss.toString();
                    rtss = null;
                } catch (NumberFormatException ex) {
                    thisStr = value.toString();
                }
                break;
            case NUMBER:
                if (formatString != null) {
                    thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
                } else {
                    thisStr = value;
                }

                thisStr = thisStr.replace("_", "").trim();
                break;
            case DATE:
                try {
                    thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
                } catch (NumberFormatException e){
                    thisStr = value.toString();
                }

                // 对日期字符串作特殊处理
                thisStr = thisStr.replace(" ", "T");
                break;
            default:
                thisStr = " ";

                break;
        }

        return thisStr;
    }

    /**
     * 解析得到列
     * @param ch
     * @param start
     * @param length
     * @throws SAXException
     */
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        // 得到单元格内容的值
        lastContents += new String(ch, start, length);
    }

    /**
     * 获取行数据回调
     * @param sheetIndex
     * @param curRow
     * @param rowMap
     */
    public abstract void getRows(int sheetIndex, int curRow, Map<String, String> rowMap);

    public String parseIndexStr(String indexStr){
        Pattern pattern = Pattern.compile("[0-9]");
        Matcher matcher = pattern.matcher(indexStr);
        String resultString = matcher
                .replaceAll("");
        return resultString;
    }

    /**
     * 测试方法
     */
    public static void main(String[] args) throws Exception {
/////////////////////////////////////////////////////////////////////////////////////////////////////////
        final int[] total = {-1};
        Map<String, String> keyMap = new HashMap<>();
        List<Map<String, String>> mapList = new ArrayList<>();

        try {
            ExcelReader reader = new ExcelReader() {
                public void getRows(int sheetIndex, int curRow, Map<String, String> rowMap) {
                    total[0]++;
                    if(curRow == 0){
                        keyMap.putAll(rowMap);
                    } else {
                        Map<String, String> map = new HashMap<>();
                        for (Map.Entry<String, String> entry: rowMap.entrySet()){
                            map.put(keyMap.get(entry.getKey()), entry.getValue());
                        }
                        mapList.add(map);
                    }
                }
            };
            reader.process("src/main/resources/22.xlsx", 1);
        } catch (Exception e) {
            e.printStackTrace();
        }

        System.out.println(mapList);
        System.out.println(total[0]);
/////////////////////////////////////////////////////////////////////////////////////////////////////////
    }

}