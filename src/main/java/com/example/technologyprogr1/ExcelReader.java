package com.example.technologyprogr1;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {
    public static List<String> readExcelFiles(List<File> files) {
        List<String> data = new ArrayList<>();
        try {
            for (File file : files) {
                try (FileInputStream fis = new FileInputStream(file);
                     Workbook workbook = new XSSFWorkbook(fis)) {
                    for (Sheet sheet : workbook) {
                        for (Row row : sheet) {
                            StringBuilder rowData = new StringBuilder();
                            for (Cell cell : row) {
                                rowData.append(cell.toString()).append(" | ");
                            }
                            data.add(rowData.toString());
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return data;
    }
}


