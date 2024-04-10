package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.BeforeEach;

import java.math.BigDecimal;
import java.util.Date;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class ExcelUtilTest {

    private Workbook workbook;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook(); // Using XSSFWorkbook for testing
   }

    @Test
    void testGetCellStyleMap() {
        List<String> formatList = Arrays.asList("0.00", "dd/mm/yyyy", "hh:mm:ss");

        Map<String, CellStyle> styleMap = ExcelUtil.getCellStyleMap(workbook, formatList);

        assertNotNull(styleMap);
        assertEquals(4, styleMap.size()); // Expecting 4 cell styles

        // Verify if the header style is correctly added
        assertTrue(styleMap.containsKey("HEADER"));
        CellStyle headerStyle = styleMap.get("HEADER");
        assertNotNull(headerStyle);
        assertEquals(new XSSFColor(new byte[]{(byte)211, (byte)211, (byte)211}), headerStyle.getFillForegroundColorColor());

        // Verify if data styles with formats are correctly added
        for (String format : formatList) {
            assertTrue(styleMap.containsKey(format));
            CellStyle dataStyle = styleMap.get(format);
            assertNotNull(dataStyle);
            assertEquals(format, dataStyle.getDataFormatString());
        }
    }

    @Test
    void testCreateSheet() {
        String sheetName = "TestSheet";
        int fromRowNum = 0;
        List<String> headers = new ArrayList<>();
        headers.add("Name");
        headers.add("Age");
        headers.add("Date of Birth");

        List<List<Object>> data = new ArrayList<>();
        data.add(new ArrayList<>(List.of("John Doe", 30, LocalDate.of(1990, 1, 1))));
        data.add(new ArrayList<>(List.of("Jane Doe", 25, LocalDate.of(1998, 2, 1))));

        List<String> cellFormatList = Arrays.asList(null, null, "yyyy-MM-dd"); 

        Map<String, CellStyle> cellStyleMap = ExcelUtil.getCellStyleMap(workbook, cellFormatList);

        // 调用待测试的方法
        ExcelUtil.createSheet(workbook, sheetName, fromRowNum, headers, data, cellFormatList, cellStyleMap);

        Sheet sheet = workbook.getSheet(sheetName);
        assertNotNull(sheet, "Sheet should be created");

        // 校验表头
        Row headerRow = sheet.getRow(fromRowNum);
        assertNotNull(headerRow);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.getCell(i);
            assertEquals(headers.get(i), cell.getStringCellValue(), "Header value should match");
        }

        // 校验数据行
        for (int i = 0; i < data.size(); i++) {
            Row dataRow = sheet.getRow(fromRowNum + i + 1);
            assertNotNull(dataRow);
            for (int j = 0; j < data.get(i).size(); j++) {
                Cell cell = dataRow.getCell(j);
                Object value = data.get(i).get(j);
                if (value instanceof BigDecimal) {
                    assertEquals(((BigDecimal) value).doubleValue(), cell.getNumericCellValue(), "BigDecimal value should match");
                } else if (value instanceof Number) {
                    assertEquals(((Number) value).doubleValue(), cell.getNumericCellValue(), "Number value should match");
                } else if (value instanceof Date) {
                    assertEquals((Date) value, cell.getDateCellValue(), "Date value should match");
                } else if (value instanceof LocalDate) {
                    //convert value to java.util.Date
                    Date date = Date.from(((LocalDate) value).atStartOfDay().atZone(java.time.ZoneId.systemDefault()).toInstant());
                    assertEquals(date, cell.getDateCellValue(), "LocalDate value should match");
                } else if (value instanceof LocalDateTime) {
                    //convert value to java.util.Date
                    Date date = Date.from(((LocalDateTime) value).atZone(java.time.ZoneId.systemDefault()).toInstant());
                    assertEquals(date, cell.getDateCellValue(), "LocalDateTime value should match");
                } else if (value instanceof String) {
                    assertEquals((String) value, cell.getStringCellValue(), "String value should match");
                }
            }
        }
    }

    
}