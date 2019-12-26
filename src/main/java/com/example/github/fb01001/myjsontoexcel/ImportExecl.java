package com.example.github.fb01001.myjsontoexcel;

import com.alibaba.fastjson.JSONObject;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;
import org.apache.poi.POIDocument;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import static org.apache.logging.log4j.util.Strings.isNotEmpty;
import static org.springframework.util.StringUtils.*;

/***
 *@Title ${TODO}
 *＠author    fangbin
 *@Date 19-12-17 上午11:04
 */
public class ImportExecl {
    /** 总行数 */
    private int totalRows = 0;
    /** 总列数 */
    private int totalCells = 0;
    /** 错误信息 */
    private String errorInfo;
    /** 构造方法 */
    public ImportExecl() {
    }


    public int getTotalRows() {
        return totalRows;
    }
    public int getTotalCells() {
        return totalCells;
    }
    public String getErrorInfo() {
        return errorInfo;
    }

    public boolean validateExcel(String filePath) {
        /** 检查文件名是否为空或者是否是Excel格式的文件 */
        if (filePath == null
                || !(this.isExcel2003(filePath) || this.isExcel2007(filePath))) {
            errorInfo = "文件名不是excel格式";
            return false;
        }
        /** 检查文件是否存在 */
        File file = new File(filePath);
        if (file == null || !file.exists()) {
            errorInfo = "文件不存在";
            return false;
        }
        return true;
    }

    public List<List<String>> read(String filePath) {
        List<List<String>> dataLst = new ArrayList<List<String>>();
        InputStream is = null;
        try {
            /** 验证文件是否合法 */
            if (!validateExcel(filePath)) {
                System.out.println(errorInfo);
                return null;
            }
            /** 判断文件的类型，是2003还是2007 */
            boolean isExcel2003 = true;
            if (this.isExcel2007(filePath)) {
                isExcel2003 = false;
            }
            /** 调用本类提供的根据流读取的方法 */
            File file = new File(filePath);
            is = new FileInputStream(file);
            dataLst = read(is, isExcel2003);
            is.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    is = null;
                    e.printStackTrace();
                }
            }
        }
        /** 返回最后读取的结果 */
        return dataLst;
    }


    public List<List<String>> read(InputStream inputStream, boolean isExcel2003) {
        List<List<String>> dataLst = null;
        try {
            /** 根据版本选择创建Workbook的方式 */
            org.apache.poi.ss.usermodel.Workbook wb = null;
            if (isExcel2003) {
                //wb = new HSSFWorkbook(inputStream);
                wb = new HSSFWorkbook(inputStream);
            } else {
                //wb = new XSSFWorkbook(inputStream);
                wb = new XSSFWorkbook(inputStream);
            }
            dataLst = read(wb);
        } catch (IOException e) {

            e.printStackTrace();
        }
        return dataLst;
    }


    private List<List<String>> read(org.apache.poi.ss.usermodel.Workbook wb) {
        List<List<String>> dataLst = new ArrayList<List<String>>();
        /** 得到第一个shell */
        org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheetAt(0);
        /** 得到Excel的行数 */
        this.totalRows = sheet.getPhysicalNumberOfRows();
                //getPhysicalNumberOfRows();
        /** 得到Excel的列数 */
        if (this.totalRows >= 1 && sheet.getRow(0) != null) {
            this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }
        /** 循环Excel的行 */
        for (int r = 0; r < this.totalRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            List<String> rowLst = new ArrayList<String>();
            /** 循环Excel的列 */
            for (int c = 0; c < row.getPhysicalNumberOfCells(); c++) {
                Cell cell = row.getCell(c);
                String cellValue = "";
                if (null != cell) {
                    // 以下是判断数据的类型
                    switch (cell.getCellTypeEnum().getCode()) {
                        case HSSFCell.CELL_TYPE_NUMERIC: // 数字
                            cellValue = cell.getNumericCellValue() + "";
                                    //getNumericCellValue() + "";
                            break;
                        case HSSFCell.CELL_TYPE_STRING: // 字符串
                            cellValue = cell.getStringCellValue();
                            break;
                        case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
                            cellValue = cell.getBooleanCellValue() + "";
                            break;
                        case HSSFCell.CELL_TYPE_FORMULA: // 公式
                            cellValue = cell.getCellFormula() + "";
                            break;
                        case HSSFCell.CELL_TYPE_BLANK: // 空值
                            cellValue = "";
                            break;
                        case HSSFCell.CELL_TYPE_ERROR: // 故障
                            cellValue = "非法字符";
                            break;
                        default:
                            cellValue = "未知类型";
                            break;
                    }
                }
                rowLst.add(cellValue);
            }
            /** 保存第r行的第c列 */
            dataLst.add(rowLst);
        }
        return dataLst;
    }
    public static void main(String[] args) throws Exception {
        ImportExecl poi = new ImportExecl();
        // List<List<String>> list = poi.read("d:/aaa.xls");
        List<List<String>> list = poi.read("/home/fangbin/Desktop/安全扫描/第二轮扫描导出结果.xlsx");
        List<List<String>> needExportSourceList = poi.read("/home/fangbin/Desktop/安全扫描/错误扫描 20191226.xlsx");
        List<String> needList = new ArrayList<String>();

        System.out.println("needExportSourceList.size()-----------" + needExportSourceList.size());
        for (List<String> strs: needExportSourceList
             ) {
            System.out.println(strs.get(0));
            if(isNotEmpty(strs.get(0))){
                needList.add(strs.get(0));
            }
        }
        System.out.println("need export source lineNum ---" + needList.size());


        if (list != null) {
            int count = 0;
            /*for (int i = 0; i < list.size(); i++) {
                List<String> cellList = list.get(i);
                if(needList.contains(cellList.get(36))){
                    count++;
                    System.out.print("第" + (i) + "行");
                    for (int j = 0; j < cellList.size(); j++) {
                        // System.out.print("    第" + (j + 1) + "列值：");
                        System.out.print("    " + cellList.get(j));
                    }
                    System.out.println();
                }
            }
            System.out.println(count);*/



            // excel 信息
            try {
                OutputStream outputStream = new FileOutputStream("/home/fangbin/Desktop/安全扫描/20191226test111.xls");// 创建工作薄
                WritableWorkbook workbook = Workbook.createWorkbook(outputStream);
                WritableWorkbook writableWorkbook = Workbook.createWorkbook(outputStream);

                // 创建新的一页
                WritableSheet sheet = workbook.createSheet("First Sheet",0);
                Label label01 = new Label(0,0,"模块");
                sheet.addCell(label01);
                Label label02 = new Label(1,0,"bug等级");
                sheet.addCell(label02);
                Label label03 = new Label(2,0,"issueInstanceId");
                sheet.addCell(label03);
                Label label04 = new Label(3,0,"fullFileName");
                sheet.addCell(label04);
                Label label05 = new Label(4,0,"issueName");
                sheet.addCell(label05);
                Label label0９ = new Label(5,0,"误报解释");
                sheet.addCell(label0９);
                Label label１０ = new Label(6,0,"lineNumber");
                sheet.addCell(label１０);

                for (int i = 0; i < list.size(); i++) {
                    List<String> cellList = list.get(i);
                    if(needList.contains(cellList.get(36))){
                        count++;
                        sheet.addCell(new Label(0,count,cellList.get(4)));//模块
                        sheet.addCell(new Label(1,count,cellList.get(7)));//ｂｕｇ等级
                        sheet.addCell(new Label(2,count,cellList.get(31)));//issueInstanceId
                        sheet.addCell(new Label(3,count,cellList.get(36)));//fullFileName
                        sheet.addCell(new Label(4,count,cellList.get(3)));//issueName
                        if("Log Forging".equalsIgnoreCase(cellList.get(3))){
                            //日志打印，误报
                            sheet.addCell(new Label(5,count,"日志打印，误报"));//issueName
                        }else if("Bean Manipulation".equalsIgnoreCase(cellList.get(5))){
                            //日志打印，误报
                            sheet.addCell(new Label(5,count,"内部参数转换，误报"));//issueName
                        }else if("SQL Injection: Hibernate".equalsIgnoreCase(cellList.get(5))){
                            //日志打印，误报
                            sheet.addCell(new Label(5,count,"数据库查询，误报"));//issueName
                        }else if("Access Control: Database".equalsIgnoreCase(cellList.get(5))){
                            //日志打印，误报
                            sheet.addCell(new Label(5,count,"数据库查询，误报"));//issueName
                        }
                        sheet.addCell(new Label(6,count,cellList.get(34).substring(0,cellList.get(34).indexOf("."))));//lineNumber
                        /*sheet.addCell(new Label(5,count,cellList.get(1)));//漏洞简述
                        sheet.addCell(new Label(6,count,cellList.get(1)));
                        sheet.addCell(new Label(7,count,cellList.get(1)));
                        sheet.addCell(new Label(8,count,cellList.get(1)));
                        sheet.addCell(new Label(9,count,cellList.get(1)));*/
                    }
                }


                //把创建的内容写入到输出流中，并关闭输出流
                workbook.write();
                workbook.close();
                outputStream.close();

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (RowsExceededException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            }
        }

    }

    public static boolean isExcel2003(String filePath) {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }
    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

}
