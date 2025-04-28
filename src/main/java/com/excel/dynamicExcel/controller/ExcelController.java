package com.excel.dynamicExcel.controller;

import com.excel.dynamicExcel.service.*;
import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;


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
        String userHome = System.getProperty("user.home");  // Example: C:/Users/Asus
        String downloadsDir = userHome + File.separator + "Downloads"; // C:/Users/Asus/Downloads

        String timestamp = new java.text.SimpleDateFormat("yyyyMMdd_HHmmss").format(new java.util.Date());
        String filename = "exported_excel_file_" + timestamp + ".xlsx";

        String outputFilePath = downloadsDir + File.separator + filename;

        excelExportService.exportDatabaseToExcel(outputFilePath);
        return ResponseEntity.ok("Excel file exported successfully to " + outputFilePath);
    }
}