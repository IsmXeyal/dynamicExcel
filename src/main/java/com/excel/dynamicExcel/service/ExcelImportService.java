package com.excel.dynamicExcel.service;

import com.excel.dynamicExcel.exception.ExcelImportException;
import com.excel.dynamicExcel.util.ExcelUtil;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.*;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
public class ExcelImportService {

    private final JdbcTemplate jdbcTemplate;

    public void importExcelToDatabase(MultipartFile file) {
        try (InputStream is = file.getInputStream(); Workbook workbook = new XSSFWorkbook(is)) {
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) throw new ExcelImportException("Sheet not found!");

            List<String> columns = ExcelUtil.getHeaderFromSheet(sheet);

            String tableName = "dynamic_excel_data";

            jdbcTemplate.execute("DROP TABLE IF EXISTS " + tableName);

            String createTableSql = "CREATE TABLE " + tableName + " ("
                    + columns.stream()
                    .map(col -> "\"" + col + "\" TEXT")
                    .collect(Collectors.joining(", "))
                    + ")";
            jdbcTemplate.execute(createTableSql);

            List<List<String>> data = ExcelUtil.getDataFromSheet(sheet, columns);

            for (List<String> values : data) {
                String insertSql = "INSERT INTO " + tableName + " ("
                        + columns.stream().map(col -> "\"" + col + "\"").collect(Collectors.joining(", "))
                        + ") VALUES ("
                        + String.join(", ", values)
                        + ")";
                jdbcTemplate.execute(insertSql);
            }
        } catch (Exception e) {
            throw new ExcelImportException("Error importing Excel file: " + e.getMessage(), e);
        }
    }
}
