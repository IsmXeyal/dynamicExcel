package com.excel.dynamicExcel.exception;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

@RestControllerAdvice
public class GlobalControllerExceptionHandler {
    @ExceptionHandler(ExcelImportException.class)
    public ResponseEntity<String> handleExcelImportException(ExcelImportException ex) {
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("Excel Import Error: " + ex.getMessage());
    }

    @ExceptionHandler(ExcelExportException.class)
    public ResponseEntity<String> handleExcelExportException(ExcelExportException ex) {
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Excel Export Error: " + ex.getMessage());
    }

    @ExceptionHandler(Exception.class)
    public ResponseEntity<String> handleGeneralException(Exception ex) {
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("An unexpected error occurred: " + ex.getMessage());
    }
}
