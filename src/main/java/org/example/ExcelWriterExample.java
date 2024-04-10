package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

public class ExcelWriterExample {

    private static final BigDecimal ONE_HUNDRED = new BigDecimal("100");
    private static final String HEADER = "header";

    private static final String CELL_FORMAT_TEXT = "@";
    private static final String CELL_FORMAT_CURRENCY = "#,#00.00;[Red](#,#00.00)";
    private static final String CELL_FORMAT_PERCENT = "0.00000%";
    private static final String CELL_FORMAT_DATE = "yyyy-MM-dd";

    public static void main(String[] args) {

        try (Workbook workbook = new XSSFWorkbook()) {
            List<String> headers = Arrays.asList("Text", "Currency", "Percent", "Integer", "Date");

            List<List<Object>> data = new ArrayList<>();
            data.add(generateRowData("text-1", new BigDecimal("12345678.90"), new BigDecimal("1.23456"), 3, LocalDate.now()));
            data.add(generateRowData("text-2", new BigDecimal("-19.98"), new BigDecimal("2.345678"), 4, LocalDate.parse("2024-04-07")));

            List<String> cellFormatList = Arrays.asList(CELL_FORMAT_TEXT, CELL_FORMAT_CURRENCY, CELL_FORMAT_PERCENT, null, CELL_FORMAT_DATE);

            Map<String, CellStyle> cellStyleMap = getCellStyleMap(workbook, cellFormatList);

            createSheet(workbook, "Sheet1", 0, headers, data, cellFormatList, cellStyleMap);

            //save workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("example.xlsx")) {
                workbook.write(fileOut);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    
    }

    private static List<Object> generateRowData(String text, BigDecimal currency, BigDecimal percent, Integer integer, LocalDate date) {
        return Arrays.asList(text, currency, divideOneHundred(percent, 7, RoundingMode.DOWN), integer, date);
    }

    private static BigDecimal divideOneHundred(BigDecimal number, int scale, RoundingMode roundingMode) {
        return number.divide(ONE_HUNDRED, scale, roundingMode);
    }

    private static Map<String, CellStyle> getCellStyleMap(Workbook workbook, List<String> formatList) {
        Map<String, CellStyle> result = new HashMap<>();

        //header
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(boldFont);

        //set cell color
        headerStyle.setFillForegroundColor(new XSSFColor(new byte[] {(byte)211, (byte)211, (byte)211}));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        result.put(HEADER, headerStyle);

        //data
        CreationHelper createHelper = workbook.getCreationHelper();
        for (String cellFormat : formatList) {
            if (cellFormat != null && !result.containsKey(cellFormat)) {
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(cellFormat));
                result.put(cellFormat, cellStyle);
            }
        }

        return result;
    }

    private static void createSheet(Workbook workbook,
                                    String sheetName,
                                    int fromRowNum,
                                    List<String> headers,
                                    List<List<Object>> data,
                                    List<String> cellFormatList,
                                    Map<String, CellStyle> cellStyleMap) {
        Sheet sheet = workbook.createSheet(sheetName); 

        //header
        Row row = sheet.createRow(fromRowNum);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(cellStyleMap.get(HEADER));
            cell.setCellValue(headers.get(i));
        }

        //data
        for (int i = 0; i < data.size(); i++) {
            row = sheet.createRow(fromRowNum + i + 1);
            for (int j = 0; j < data.get(i).size(); j++) {
                Cell cell = row.createCell(j);

                if (cellFormatList.get(j) != null) {
                    cell.setCellStyle(cellStyleMap.get(cellFormatList.get(j)));
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
                } else if (obj != null) {
                    cell.setCellValue(obj.toString());
                }
            }
        }

        for (int i = 0; i < headers.size(); i++) {
            sheet.autoSizeColumn(i);
        }
    }
}
