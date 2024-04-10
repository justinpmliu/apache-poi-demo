package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.util.*;

public class ExcelWriterExample {

    private static final BigDecimal ONE_HUNDRED = new BigDecimal("100");

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

            Map<String, CellStyle> cellStyleMap = ExcelUtil.getCellStyleMap(workbook, cellFormatList);

            ExcelUtil.createSheet(workbook, "Sheet1", 0, headers, data, cellFormatList, cellStyleMap);

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
}
