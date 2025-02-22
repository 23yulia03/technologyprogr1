package com.example.technologyprogr1;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class ExcelWriter {
    public static void writeToExcel(List<String> data, File outputFile) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream(outputFile)) {
            Sheet sheet = workbook.createSheet("Данные");

            Set<String> headerSet = new HashSet<>();  // Множество для отслеживания заголовков

            int rowIndex = 0;
            for (String rowText : data) {
                Row row = sheet.createRow(rowIndex++);
                String[] cells = rowText.split(" \\| ");

                // Проверяем, является ли строка заголовком и уникален ли он
                if (rowIndex == 1 || !headerSet.contains(cells[0])) {
                    for (int i = 0; i < cells.length; i++) {
                        row.createCell(i).setCellValue(cells[i]);
                    }
                    if (rowIndex == 1) {
                        // Добавляем заголовки в множество
                        for (String cell : cells) {
                            headerSet.add(cell);
                        }
                    }
                } else {
                    // Строка является дублирующимся заголовком, пропускаем
                    rowIndex--;
                }
            }

            workbook.write(out);
        }
    }
}
