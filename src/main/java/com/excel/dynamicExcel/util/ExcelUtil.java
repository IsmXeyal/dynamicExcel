package com.excel.dynamicExcel.util;

import org.apache.poi.ss.usermodel.*;

import java.time.ZoneId;
import java.util.*;

public class ExcelUtil {

    // Excel faylından bashlıq setirini oxumaq
    public static List<String> getHeaderFromSheet(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        List<String> columns = new ArrayList<>();

        for (Cell cell : headerRow) {
            String columnName = switch (cell.getCellType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                default -> "Unknown";
            };
            columns.add(columnName.trim());
        }

        return columns;
    }

    // Read data rows
    public static List<List<String>> getDataFromSheet(Sheet sheet, List<String> columns) {
        List<List<String>> data = new ArrayList<>();

        // Burada sheet.getLastRowNum() Excel sehifesinin sonuncu setirinin indeksini alır. Setirlər sıfırdan bashlayır,
        // ona gore i=1 olaraq bashlanır (bu, bashlıq setirini atlamaq uchundur).
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row dataRow = sheet.getRow(i);
            if (dataRow == null) continue;

            List<String> rowValues = new ArrayList<>();
            for (int j = 0; j < columns.size(); j++) {
                Cell cell = dataRow.getCell(j);
                String cellValue = null;

                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            cellValue = cell.getStringCellValue().replace("'", "''");
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                cellValue = "'" + cell.getDateCellValue().toInstant()
                                        .atZone(ZoneId.systemDefault())
                                        .toLocalDate()
                                        .toString() + "'";
                            } else {
                                double numericValue = cell.getNumericCellValue();
                                if (numericValue == (int) numericValue) {
                                    cellValue = String.valueOf((int) numericValue);
                                } else {
                                    cellValue = String.format("%f", numericValue);
                                }
                            }
                            break;
                        case BOOLEAN:
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            cellValue = cell.getCellFormula();
                            break;
                        case BLANK:
                            cellValue = null;
                            break;
                        default:
                            break;
                    }
                }

                if (cellValue == null) {
                    rowValues.add("null");
                } else {
                    rowValues.add("'" + cellValue + "'");
                }
            }
            data.add(rowValues);
        }
        return data;
    }
}
