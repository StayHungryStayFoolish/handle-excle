package com.excel.handle.handleexcle.test;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @Author: Created by bonismo@hotmail.com on 2019-12-13 17:40
 * @Description:
 * @Version: 1.0
 */
public class HandleExcel {

    private final static String XLS = "xls";
    private final static String XLSX = "xlsx";

    public static final String FILE_NAME = "/Users/bonismo/Downloads/新建文件夹/插入文件.xls";


    public static void showExcel() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File(FILE_NAME)));

        HSSFSheet sheet = null;
        //获取每个Sheet表
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            sheet = workbook.getSheetAt(i);
            //获取每行
            for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {
                HSSFRow row = sheet.getRow(j);
                //获取每个单元格
                for (int k = 0; k < row.getPhysicalNumberOfCells(); k++) {
                    System.out.print(row.getCell(k) + "\t");
                }
                System.out.println("---Sheet表" + i + "处理完毕---");
            }
        }
    }

    public static void commonShowExcel(String filePath) {
        Workbook workbook = getWorkBook(filePath);
        Sheet sheet = null;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            sheet = getSheet(filePath, workbook, i);
            String sheetName = sheet.getSheetName();
            System.out.println("Sheet Name : " + sheetName);
        }
    }

    public static void getSheetName() throws IOException {
        XSSFWorkbook xssfWorkBook = new XSSFWorkbook(new FileInputStream(FILE_NAME));
        XSSFSheet xssfSheet = null;
        for (int i = 0; i < xssfWorkBook.getNumberOfSheets(); i++) {
            xssfSheet = xssfWorkBook.getSheetAt(i);
            //sheet名称，用于校验模板是否正确
            String sheetName = xssfSheet.getSheetName();

        }
    }

    public static Workbook getWorkBook(String filePath) {
        File file = new File(filePath);
        //获得文件名
        String fileName = file.getName();
        String fileType = fileName.substring(fileName.lastIndexOf("."), fileName.length());
        //创建Workbook工作薄对象，表示整个excel
        Workbook workbook = null;
        try {
            //获取excel文件的io流
            InputStream is = new FileInputStream(file);
            //根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
            if(fileType.endsWith(XLS)){
                //2003
                workbook = new HSSFWorkbook(is);
            }else if(fileType.endsWith(XLSX)){
                //2007
                workbook = new XSSFWorkbook(is);
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        return workbook;
    }

    public static Sheet getSheet(String filePath,Workbook workbook,int i) {
        File file = new File(filePath);
        //获得文件名
        String fileName = file.getName();
        String fileType = fileName.substring(fileName.lastIndexOf("."), fileName.length());
        //创建Workbook工作薄对象，表示整个excel
        Sheet sheet = null;
        try {
            //获取excel文件的io流
            InputStream is = new FileInputStream(file);
            //根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
            if(fileType.endsWith(XLS)){
                //2003
                sheet = workbook.getSheetAt(i);
            }else if(fileType.endsWith(XLSX)){
                //2007
                sheet = workbook.getSheetAt(i);
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        return sheet;
    }

    public static void main(String[] args) {
        commonShowExcel(FILE_NAME);
    }
}
