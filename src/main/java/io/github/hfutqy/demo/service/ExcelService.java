package io.github.hfutqy.demo.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author qiyu
 * @date 2020/3/18
 */
public class ExcelService {


    public static void main(String[] args) throws Exception{
        File file = new File("D:\\test\\from.xlsx");
        InputStream inputStream = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(inputStream);
        //获得当前sheet工作表
        Sheet sheet = workbook.getSheetAt(0);
        //获得当前sheet的结束行
        int lastRowNum = sheet.getLastRowNum();
        // 提取横向 元素
        List<String> rowNameList = new ArrayList<>();
        Row firstRow = sheet.getRow(0);
        int firstRowCellNum = firstRow.getLastCellNum();
        for (int i = 1; i <= firstRowCellNum; i++) {
            Cell cell = firstRow.getCell(i);
            if (cell != null) {
                rowNameList.add(cell.getStringCellValue());
            }
        }
        // 提取纵向 行为
        List<String> columnNameList = new ArrayList<>();
        for (int i = 1; i <= lastRowNum; i++) {
            Cell cell = sheet.getRow(i).getCell(0);
            if (cell != null) {
                columnNameList.add(cell.getStringCellValue());
            }
        }
        System.out.println(rowNameList);
        System.out.println(columnNameList);
        List<List<String>> result = new ArrayList<>();
        //循环除了所有行,如果要循环除第一行以外的就firstRowNum+1
        for (int rowNum = 1; rowNum <= lastRowNum; rowNum++) {
            List<String> rowHit = new ArrayList<>();
            //获得当前行
            Row row = sheet.getRow(rowNum);
            //获得当前行的列数
            int lastCellNum = row.getLastCellNum();
            for (int cellNum = 1; cellNum <= lastCellNum; cellNum++) {
                Cell cell = row.getCell(cellNum);
                if (cell != null) {
                    rowHit.add(rowNameList.get(cellNum-1));
                }
            }
            result.add(rowHit);
        }
        System.out.println(result);

        // 开始输出excel
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        Sheet newSheet = xssfWorkbook.createSheet("Sheet1");
        // 行数计算
        int rowNum = 0;
        for (List<String> strings : result) {
            rowNum += strings.size();
        }
        System.out.println(rowNum);

        int allRowNum = 0;
        for (int i = 0; i < result.size(); i++) {
            List<String> secNameList = result.get(i);
            String firstName = columnNameList.get(i);
            for (String secName : secNameList) {
                Row newRow = newSheet.createRow(allRowNum);
                Cell cell0 = newRow.createCell(0);
                cell0.setCellValue(firstName);
                Cell cell1 = newRow.createCell(1);
                cell1.setCellValue(secName);
                allRowNum++;
            }
        }

        FileOutputStream output = new FileOutputStream("D:\\test\\to.xlsx");
        xssfWorkbook.write(output);
        output.flush();

    }

}
