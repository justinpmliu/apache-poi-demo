package org.example;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class ExcelUtil {
    private static final String HEADER = "HEADER";
    
    private ExcelUtil() {}

    public static Map<String, CellStyle> getCellStyleMap(Workbook workbook, List<String> formatList) {
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

    public static void createSheet(Workbook workbook,
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

                if (cellFormatList.size() > j && cellFormatList.get(j) != null) {
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
