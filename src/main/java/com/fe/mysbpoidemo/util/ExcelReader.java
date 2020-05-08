package com.fe.mysbpoidemo.util;

import com.fe.mysbpoidemo.model.CellData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;
import java.net.URLEncoder;
import java.text.*;
import java.util.*;
import java.util.logging.Logger;

/**
 * 用于读取Excel文件的类
 * 解析步骤：获取workbook -> 获取sheet -> 遍历row -> 遍历cell -> 获取cell value
 * 【注】
 * （1）读取xlsx需要commons-collections4和commons-compress
 * （2）读取xls需要commons-math3
 * （3）poi-ooxml是ooxml-schemas的精简版，poi-3.7以上需要另外导入ooxml-schemas-1.1
 */
public class ExcelReader {

    private static Logger logger = Logger.getLogger(ExcelReader.class.getName()); // 日志打印类

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";

    /**
     * 根据文件后缀获取对应的工作簿对象
     *
     * @param inputStream 输入流，用于读取文件
     * @param fileType    文件后缀名类型（xls或xlsx）
     * @return Workbook对象
     * @throws IOException
     */
    public static Workbook getWorkbook(InputStream inputStream, String fileType) throws IOException {
        Workbook workbook = null;
        // 忽略大小写
        if (fileType.equalsIgnoreCase(XLS)) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (fileType.equalsIgnoreCase(XLSX)) {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }

    /**
     * 将单元格内容转换为字符串
     * （建议在遍历每一行内使用）
     *
     * @param cell 单元格
     * @return
     * @author rk
     */
    private static String convertCellValueToString(Cell cell) {
        if (cell == null) {
            return null;
        }
        //String dataFormat = cell.getCellStyle().getDataFormatString();
        //System.out.print("格式：" + dataFormat + "，类型：" + cell.getCellType() + " ");
        String returnValue = null;
        switch (cell.getCellType()) {
            case NUMERIC:   //数字
                // 获取单元格格式
                String dataFormat = cell.getCellStyle().getDataFormatString();
                /*
                 * 单元格格式
                 * General： 常规
                 * @：文本
                 * 0_：数值，0位小数
                 * 0.0_：数值：1位小数
                 * ......
                 */
                if (dataFormat.equals("General") || dataFormat.equals("@") || dataFormat.contains("0_")) {
                    /*
                    // 直接将double转为string不会保留0，如果没有小数也会带个.0
                    returnValue = String.valueOf(cell.getNumericCellValue());
                    */
                    if (dataFormat.contains("0_")) {
                        // 保留多少位数
                        // int decimalBit = dataFormat.substring(dataFormat.indexOf(".") + 1, dataFormat.indexOf("_")).length();
                        Double doubleValue = cell.getNumericCellValue();
                        // 数值格式
                        String numFormatStr = dataFormat.substring(0, dataFormat.indexOf("_"));
                        /*
                         * 格式化科学计数法，例如
                         * 传入0.00表示保留2位小数，如123.4 -> 123.40
                         * 传入0.000表示保留3位小数，例如1234 -> 1234.000
                         */
                        DecimalFormat df = new DecimalFormat(numFormatStr);
                        returnValue = df.format(doubleValue);
                    } else {
                        Double doubleValue = cell.getNumericCellValue();
                        /*
                         * 格式化科学计数法，取整数
                         * 如果采用常规格式保存小数则会丢弃小数位，例如123.4 -> 123
                         */
                        DecimalFormat df = new DecimalFormat("0");
                        returnValue = df.format(doubleValue);
                    }
                } else {
                    // 处理日期
                    returnValue = convertDateToString(dataFormat, cell.getDateCellValue());
                }
                break;
            case STRING:    //字符串
                returnValue = cell.getStringCellValue();
                break;
            case BOOLEAN:   //布尔
                Boolean booleanValue = cell.getBooleanCellValue();
                returnValue = booleanValue.toString();
                break;
            case BLANK:     // 空值
                returnValue = null;
                break;
            case FORMULA:   // 公式
                returnValue = cell.getCellFormula();
                break;
            case ERROR:     // 故障
                returnValue = "[ERROR]";
                break;
            default:
                break;
        }
        return returnValue;
    }

    /**
     * 将数字转换为日期或时间
     * 【注】POI对单元格日期没有针对的类型，日期类型取出来的也是一个double值(NUMERIC)
     *
     * @param cellDataFormatString 通过cell.getCellStyle().getDataFormatString()得到的String
     * @param cellDateValue        通过cell.getDateCellValue()得到的Date值
     * @return
     * @author rk
     */
    private static String convertDateToString(String cellDataFormatString, Date cellDateValue) {
        System.out.println("单元格格式：" + cellDataFormatString);
        String dateValue;
        // 默认日期格式
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        switch (cellDataFormatString) {
            default:
                dateValue = simpleDateFormat.format(cellDateValue);
                break;
            case "m/d/yy":
                dateValue = new SimpleDateFormat("yyyy/M/d").format(cellDateValue);
                break;
            case "yyyy/mm/dd":
                dateValue = new SimpleDateFormat("yyyy/MM/dd").format(cellDateValue);
                break;
            case "yyyy\\-mm\\-dd":
                dateValue = simpleDateFormat.format(cellDateValue);
                break;
            case "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy":
                dateValue = new SimpleDateFormat("yyyy年M月d日").format(cellDateValue);
                break;
            case "yyyy\"年\"mm\"月\"dd\"日\"":
                dateValue = new SimpleDateFormat("yyyy年MM月dd日").format(cellDateValue);
                break;
            case "yyyy\\-mm\\-dd\\ hh:mm:ss":
                dateValue = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(cellDateValue);
                break;
            case "yyyy\\-m\\-d\\ hh:mm:ss":
                dateValue = new SimpleDateFormat("yyyy-M-d HH:mm:ss").format(cellDateValue);
                break;
            case "yyyy/mm/dd hh:mm:ss":
                dateValue = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").format(cellDateValue);
                break;
            case "yyyy/m/d hh:mm:ss":
                dateValue = new SimpleDateFormat("yyyy/M/d HH:mm:ss").format(cellDateValue);
                break;
            case "mm/dd/yyyy":
                dateValue = new SimpleDateFormat("MM/dd/yyyy").format(cellDateValue);
                break;
        }
        return dateValue;
    }

    /**
     * 遍历表格单元格（用于测试POI）（已废弃）
     *
     * @param wb 工作簿
     */
    @Deprecated
    private static void readWorkbook(Workbook wb) {
        List<Sheet> sheetList = new ArrayList<>(); // 存放所有sheet
        // 遍历并添加所有sheet
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            sheetList.add(wb.getSheetAt(i));
        }
        // 遍历每个sheet
        for (int j = 0; j < sheetList.size(); j++) {
            Sheet sheet = wb.getSheetAt(j);
            for (Iterator ite = sheet.rowIterator(); ite.hasNext(); ) {
                // 每一行
                Row row = (Row) ite.next();
                System.out.println();
                for (Iterator itet = row.cellIterator(); itet.hasNext(); ) {
                    // 每一个单元格
                    Cell cell = (Cell) itet.next();
                    // 打印单元格内容
                    System.out.print(convertCellValueToString(cell) + " ");
                }
            }
            System.out.println();
        }
    }

    /**
     * （测试POI读取excel）输出每个单元格的内容（已废弃）
     *
     * @param fileName excel文件路径
     * @author rk
     */
    @Deprecated
    public static void readExcel(String fileName) {
        Workbook workbook = null;
        FileInputStream inputStream = null;
        try {
            // 获取Excel后缀名
            String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
            // 获取Excel文件
            File excelFile = new File(fileName);
            if (!excelFile.exists()) {
                logger.warning("指定的Excel文件不存在！");
            }
            // 获取Excel工作簿
            inputStream = new FileInputStream(excelFile);
            workbook = getWorkbook(inputStream, fileType);
            readWorkbook(workbook);
        } catch (Exception e) {
            logger.warning("解析Excel失败，文件名：" + fileName + " 错误信息：" + e.getMessage());
            //e.printStackTrace();
        } finally {
            try {
                if (null != workbook) {
                    workbook.close();
                }
                if (null != inputStream) {
                    inputStream.close();
                }
            } catch (Exception e) {
                logger.warning("关闭数据流出错！错误信息：" + e.getMessage());
            }
        }
    }

    /**
     * 获取工作簿的第一个表格
     * （在线读取，https转http）
     *
     * @param filePath 文件路径
     * @return 表格对象
     * @throws IOException
     */
    public static Sheet getFirstSheet(String filePath) throws IOException {
        // 获取文件后缀
        String fileSuffix = FileUtil.getFileSuffix(filePath);
        // 在线读取（不用下载到本地） 将https转为http
        InputStream inputStream = new URL(FileUtil.solveHttps(filePath)).openStream();
        //Excel工作簿
        //Workbook wb = WorkbookFactory.create(inputStream);
        Workbook wb = getWorkbook(inputStream, fileSuffix);
        //sheet表格
        Sheet sheet = wb.getSheetAt(0);
        return sheet;
    }

    /**
     * 获取表格内第一行的标题属性（已废弃）
     *
     * @param sheet 表格
     * @return String[]
     */
    @Deprecated
    public static String[] getHeadNames(Sheet sheet) {
        String[] headNames = null;
        if (sheet != null) {
            Row row = sheet.getRow(0);
            headNames = new String[row.getPhysicalNumberOfCells()];
            //每列
            for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                Cell cell = row.getCell(j); // 第j列
                // 列不为空
                if (cell != null) {
                    String cellValue = convertCellValueToString(cell);
                    if (cellValue != null) {
                        headNames[j] = cellValue;
                    } else {
                        headNames[j] = "null";
                    }
                }
            } // for end 获取标题结束
        }
        return headNames;
    }

    /**
     * 处理空列
     * 2020-05-08 11:30
     *
     * @param sheet
     * @return
     */
    public static Sheet solveNullCell(Sheet sheet) {
        // 第1行为属性（属性为空则丢弃该列）
        Row firstRow = sheet.getRow(0);
        // 总列数
        int totalCell = firstRow.getLastCellNum();
        System.out.println("第1行总列数：" + totalCell);
        // 遍历第1行每一列
        for (int i = 0; i < totalCell; i++) {
            // 当前列
            Cell cell = firstRow.getCell(i);
            // 列判空
            if (cell == null) {
                // System.out.println("第" + (i + 1) + "列为空");
                // 当前列为空，将当前列和后面的列整体左移1格
                sheet.shiftColumns(i + 1, firstRow.getLastCellNum(), -1); // 整列左移
                // firstRow.shiftCellsLeft(i + 1, firstRow.getLastCellNum(), -1); // 行内列左移会报错
                // 总列数-1
                totalCell--;
                // 当前列-1
                i--;
                continue;
            }
        }
        return sheet;
    }

    /**
     * 处理空行
     * 2020-05-08 12:15
     *
     * @param sheet
     * @return
     */
    public static Sheet solveNullRow(Sheet sheet) {
        // 总行数
        int totalRow = sheet.getLastRowNum() + 1;
        System.out.println("总行数：" + totalRow);
        // 遍历每一行
        for (int i = 0; i < totalRow; i++) {
            // 当前行
            Row row = sheet.getRow(i);
            // 行判空
            if (row == null) {
                System.out.println("第" + (i + 1) + "行为空");
                // 如果是空行（即没有任何数据、格式），直接把它以下的数据往上移动
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
                // 总行数-1
                totalRow--;
                // 当前行-1
                i--;
                continue;
            }
        }
        return sheet;
    }

    /**
     * 获取表格内第一行的标题属性
     * （列单元格为null的丢弃）
     *
     * @param sheet 表格
     * @return
     */
    public static List<String> getHeadNameList(Sheet sheet) {
        List<String> headNameList = new ArrayList<>();
        if (sheet != null) {
            // 第1行的视为属性
            Row row = sheet.getRow(0);
            // System.out.println("总共多少列属性："+row.getPhysicalNumberOfCells()); // 返回数不含空列
            // System.out.println("实际多少列："+row.getLastCellNum());
            // 列数
            int lastCellNum = row.getLastCellNum();
            // 遍历每列
            for (int j = 0; j < lastCellNum; j++) {
                Cell cell = row.getCell(j); // 第j列
                // 列不为空
                if (cell != null) {
                    String cellValue = convertCellValueToString(cell);
                    if (cellValue != null) {
                        headNameList.add(cellValue);
                    }
                }
            } // for end 获取标题结束
        }
        return headNameList;
    }

    /**
     * 获取单元格内容（String），建议在遍历每一行时传入参数
     *
     * @param cell 单元格
     * @return
     */
    public static String getCellValue(Cell cell) {
        return convertCellValueToString(cell);
    }

    /**
     * 将数据封装为JSON格式
     * （已废弃，浏览器返回的数据不合适）
     *
     * @param filePath 文件路径
     * @return
     */
    @Deprecated
    public static List<String> readExcelToJson(String filePath) {
        List<String> dataList = new ArrayList<>();
        try {
            Sheet sheet = getFirstSheet(filePath);
            String[] headNames = getHeadNames(sheet);
            if (sheet != null) {
                // 从第2行开始遍历每一行，组装对象
                for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                    Row row = sheet.getRow(i); // 第i行
                    // 当前行不为空
                    if (row != null) {
                        StringBuilder stringBuilder = new StringBuilder(); // 临时生成Json数据
                        stringBuilder.append("{");
                        // 遍历第i行每一列
                        for (int j = 0; j < headNames.length; j++) {
                            Cell cell = row.getCell(j); // 当前单元格
                            String value = getCellValue(cell);
                            //System.out.println("第" + (i+1) +"行，第" + (j + 1) + "列，值：" + value);

                            if (j != headNames.length - 1) {
                                // 如果不是第i行的最后一列，则添加","
                                stringBuilder.append("\"").append(headNames[j]).append("\"").append(":").append("\"").append(value).append("\"").append(",");
                            } else {
                                stringBuilder.append("\"").append(headNames[j]).append("\"").append(":").append("\"").append(value).append("\"");
                            }
                        } // 遍历每一列结束
                        if (i != sheet.getPhysicalNumberOfRows() - 1) {
                            stringBuilder.append("},");
                        } else {
                            stringBuilder.append("}");
                        }
                        dataList.add(stringBuilder.toString());
                    }
                } // 遍历每一行结束
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return dataList;
    }

}
