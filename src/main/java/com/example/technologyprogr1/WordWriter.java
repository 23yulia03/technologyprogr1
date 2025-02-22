package com.example.technologyprogr1;

import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class WordWriter {
    public static void writeToWord(List<String> data, File outputFile) throws IOException {
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(outputFile)) {

            // Определяем количество колонок по первой строке
            int columnCount = data.isEmpty() ? 1 : data.get(0).split(" \\| ").length;

            // Создаем таблицу
            XWPFTable table = document.createTable();

            // Если таблица пустая, удалим первую строку, которая создается по умолчанию
            if (table.getRows().size() > 0) {
                table.removeRow(0);
            }

            Set<String> headerSet = new HashSet<>();  // Множество для отслеживания заголовков

            // Заполняем таблицу данными
            for (String rowText : data) {
                XWPFTableRow tableRow = table.createRow();
                String[] cells = rowText.split(" \\| ");

                // Проверяем, является ли строка заголовком и уникален ли он
                if (table.getRows().size() == 1 || !headerSet.contains(cells[0])) {
                    for (int i = 0; i < columnCount; i++) {
                        XWPFTableCell cell = i < tableRow.getTableCells().size()
                                ? tableRow.getCell(i)
                                : tableRow.addNewTableCell();
                        cell.setText(i < cells.length ? cells[i] : ""); // Заполняем ячейки
                    }
                    if (table.getRows().size() == 1) {
                        // Добавляем заголовки в множество
                        for (String cell : cells) {
                            headerSet.add(cell);
                        }
                    }
                } else {
                    // Пропускаем строку с дублирующимся заголовком
                    table.removeRow(table.getRows().size() - 1);
                }
            }

            document.write(out);
        }
    }
}
