package com.example.technologyprogr1;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelWriter {
    public static void writeToExcel(List<String> data, File outputFile) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream(outputFile)) {
            Sheet sheet = workbook.createSheet("Данные");

            int rowIndex = 0;
            for (String rowText : data) {
                Row row = sheet.createRow(rowIndex++);
                String[] cells = rowText.split(" \\| ");
                for (int i = 0; i < cells.length; i++) {
                    row.createCell(i).setCellValue(cells[i]);
                }
            }

            workbook.write(out);
        }
    }
}
