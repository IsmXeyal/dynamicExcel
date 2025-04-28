package com.excel.dynamicExcel.exception;

public class ExcelImportException extends RuntimeException {
    public ExcelImportException(String message) {
        super(message);
    }

    public ExcelImportException(String message, Throwable cause) {
        super(message, cause);
    }
}
