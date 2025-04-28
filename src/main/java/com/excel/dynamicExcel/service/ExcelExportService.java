package com.excel.dynamicExcel.service;

import com.excel.dynamicExcel.exception.ExcelExportException;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;

@Service
@RequiredArgsConstructor
public class ExcelExportService {
    private final JdbcTemplate jdbcTemplate;

    public void exportDatabaseToExcel(String filePath) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("ExportedData");

        List<Map<String, Object>> rows = jdbcTemplate.queryForList("SELECT * FROM dynamic_excel_data");

        if (rows.isEmpty()) {
            try {
                workbook.close();
            } catch (IOException ignored) {}
            throw new ExcelExportException("No data to export.");
        }

        Set<String> columns = rows.get(0).keySet();
        int rowIdx = 0;
        Row headerRow = sheet.createRow(rowIdx++);

        Font boldFont = workbook.createFont();
        boldFont.setBold(true);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(boldFont);

        int cellIdx = 0;
        for (String column : columns) {
            Cell cell = headerRow.createCell(cellIdx++);
            cell.setCellValue(column);
            cell.setCellStyle(headerStyle);
        }

        for (Map<String, Object> rowData : rows) {
            Row dataRow = sheet.createRow(rowIdx++);
            cellIdx = 0;
            for (String column : columns) {
                Cell cell = dataRow.createCell(cellIdx++);
                Object value = rowData.get(column);
                cell.setCellValue(value != null ? value.toString() : "");
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
        } catch (IOException e) {
            throw new ExcelExportException("Error exporting Excel file: " + e.getMessage(), e);
        } finally {
            try {
                workbook.close();
            } catch (IOException ignored) {}
        }
    }
}
