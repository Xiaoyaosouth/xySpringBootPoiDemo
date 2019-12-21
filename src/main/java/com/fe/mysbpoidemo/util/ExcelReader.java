package com.fe.mysbpoidemo.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
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
     * 将单元格内容转换为字符串（建议在遍历每一行内使用）
     *
     * @param cell 单元格
     * @return
     */
    private static String convertCellValueToString(Cell cell) {
        if (cell == null) {
            return null;
        }
        String returnValue = null;
        switch (cell.getCellType()) {
            case NUMERIC:   //数字
                Double doubleValue = cell.getNumericCellValue();
                // 格式化科学计数法，取一位整数
                DecimalFormat df = new DecimalFormat("0");
                returnValue = df.format(doubleValue);
                break;
            case STRING:    //字符串
                returnValue = cell.getStringCellValue();
                break;
            case BOOLEAN:   //布尔
                Boolean booleanValue = cell.getBooleanCellValue();
                returnValue = booleanValue.toString();
                break;
            case BLANK:     // 空值
                break;
            case FORMULA:   // 公式
                returnValue = cell.getCellFormula();
                break;
            case ERROR:     // 故障
                break;
            default:
                break;
        }
        return returnValue;
    }

    /**
     * 遍历表格单元格（用于测试POI）
     *
     * @param wb 工作簿
     */
    private static void readWorkbook(Workbook wb) {
        List<Sheet> sheetList = new ArrayList<>(); // 存放所有sheet
        // 遍历并添加所有sheet
        for (int i = 0; i < wb.getNumberOfSheets(); i++){
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
     * （测试POI读取excel）输出每个单元格的内容
     * @author rk
     * @param fileName excel文件路径
     */
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
     * @param filePath 文件路径
     * @return 表格对象
     * @throws IOException
     */
    public static Sheet getFirstSheet(String filePath) throws IOException {
        // 获取文件后缀
        String fileSuffix = FileUtil.getFileSuffix(filePath);
        // 在线读取（不用下载到本地）
        InputStream inputStream = new URL(FileUtil.solveHttps(filePath)).openStream();
        //Excel工作簿
        //Workbook wb = WorkbookFactory.create(inputStream);
        Workbook wb = getWorkbook(inputStream, fileSuffix);
        //sheet表格
        Sheet sheet = wb.getSheetAt(0);
        return sheet;
    }

    /**
     * 获取表格内第一行的标题属性
     *
     * @param sheet 表格
     * @return String[]
     */
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
     * 获取表格内第一行的标题属性
     * （列单元格为null的丢弃）
     * @param sheet 表格
     * @return
     */
    public static List<String> getHeadNameList(Sheet sheet) {
        List<String> headNameList = new ArrayList<>();
        if (sheet != null) {
            Row row = sheet.getRow(0);
            //每列
            for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
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
     * @param cell 单元格
     * @return
     */
    public static String getCellValue(Cell cell) {
        return convertCellValueToString(cell);
    }

    public static List<String> readExcelToJson(String filePath){
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
                        }else {
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
