package com.excel.dynamicExcel.controller;

import com.excel.dynamicExcel.service.*;
import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;


@RestController
@RequiredArgsConstructor
@RequestMapping("/excel")
public class ExcelController {

    private final ExcelImportService excelImportService;
    private final ExcelExportService excelExportService;

    @PostMapping("/import")
    public ResponseEntity<String> importExcel(@RequestParam("file") MultipartFile file) {
        excelImportService.importExcelToDatabase(file);
        return ResponseEntity.ok("Excel file imported successfully.");
    }

    @GetMapping("/export")
    public ResponseEntity<String> exportExcel() {
        String outputFilePath = "exported_excel_file.xlsx";
        excelExportService.exportDatabaseToExcel(outputFilePath);
        return ResponseEntity.ok("Excel file exported successfully to " + outputFilePath);
    }
}