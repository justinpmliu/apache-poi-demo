package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

public class Main {

    private static BigDecimal ONE_HUNDRED = new BigDecimal("100");

    public static void main(String[] args) throws IOException {
        Workbook workbook = new XSSFWorkbook(); // 创建新的Excel工作簿

        List<String> headers = Arrays.asList("Text", "Currency", "Percent", "Integer", "Date");

        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList("text-1", new BigDecimal("12345678.90"),
                divideOneHundred(new BigDecimal("0.01"), 12, RoundingMode.DOWN), 3, new Date()));
        data.add(Arrays.asList("text-2", new BigDecimal("12345678.99"),
                divideOneHundred(new BigDecimal("0.02"), 12, RoundingMode.DOWN), 4, LocalDate.parse("2024-04-06")));

        List<String> cellFormat = Arrays.asList(null, "#,##0.00", "0.000000000000%", null, "yyyy-MM-dd");
        createSheet(workbook, "Sheet1", headers, data, cellFormat);

        // 写入文件
        FileOutputStream fileOut = new FileOutputStream("example.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        workbook.close();
    }

    private static BigDecimal divideOneHundred(BigDecimal number, int scale, RoundingMode roundingMode) {
        return number.divide(ONE_HUNDRED, scale, roundingMode);
    }

    private static void createSheet(Workbook workbook, String sheetName, List<String> headers, List<List<Object>> data, List<String> cellFormat) {
        Sheet sheet = workbook.createSheet(sheetName); // 创建一个工作表

        //header
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(boldFont);

        Row row = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(headers.get(i));
        }

        //data
        CreationHelper createHelper = workbook.getCreationHelper();
        for (int i = 0; i < data.size(); i++) {
            row = sheet.createRow(i + 1);
            for (int j = 0; j < data.get(i).size(); j++) {
                Cell cell = row.createCell(j);

                if (cellFormat.get(j) != null) {
                    CellStyle cellStyle = workbook.createCellStyle();
                    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(cellFormat.get(j)));
                    cell.setCellStyle(cellStyle);
                }

                Object obj = data.get(i).get(j);
                if (obj instanceof BigDecimal bigDecimal) {
                    cell.setCellValue(bigDecimal.doubleValue());
                } else if (obj instanceof Number number) {
                    cell.setCellValue(number.doubleValue());
                } else if (obj instanceof Date date) {
                    cell.setCellValue(date);
                } else if (obj instanceof LocalDate localDate) {
                    cell.setCellValue(localDate);
                } else if (obj instanceof LocalDateTime localDateTime) {
                    cell.setCellValue(localDateTime);
                } else if (obj instanceof String str) {
                    cell.setCellValue(str);
                }
            }
        }

        for (int i = 0; i < headers.size(); i++) {
            sheet.autoSizeColumn(i);
        }
    }
}
