package com.example.technologyprogr1;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.Label;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.awt.*;
import java.io.*;
import java.util.List;
import java.util.Optional;

public class HelloController {
    @FXML
    private Label fileListLabel;
    @FXML
    private Label folderPathLabel;

    private StringBuilder fileList = new StringBuilder();
    private File outputFile;
    private Stage stage;

    public void setStage(Stage stage) {
        this.stage = stage;
    }

    @FXML
    protected void onSelectFiles() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
        List<File> files = fileChooser.showOpenMultipleDialog(stage);
        if (files != null) {
            fileList.setLength(0);
            for (File file : files) {
                fileList.append(file.getAbsolutePath()).append("\n");
            }
            fileListLabel.setText("Выбранные файлы:\n" + fileList.toString());
        }
    }

    @FXML
    protected void onSelectFolder() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Выберите место сохранения и имя файла");
        fileChooser.setInitialFileName("Результат.xlsx");
        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Excel Files (*.xlsx)", "*.xlsx"),
                new FileChooser.ExtensionFilter("Word Documents (*.docx)", "*.docx")
        );

        outputFile = fileChooser.showSaveDialog(stage);
        if (outputFile != null) {
            folderPathLabel.setText("Файл будет сохранен как: " + outputFile.getAbsolutePath());
        }
    }

    @FXML
    protected void onConvert() {
        if (fileList.length() == 0 || outputFile == null) {
            showAlert("Ошибка", "Выберите файлы и укажите имя файла для сохранения");
            return;
        }

        List<String> choices = List.of("Excel → Excel", "Excel → Word");
        ChoiceDialog<String> dialog = new ChoiceDialog<>(choices.get(0), choices);
        dialog.setTitle("Выбор формата");
        dialog.setHeaderText("Выберите формат итогового файла:");
        dialog.setContentText("Формат:");

        Optional<String> result = dialog.showAndWait();
        if (result.isEmpty()) return;

        try {
            if ("Excel → Excel".equals(result.get())) {
                mergeExcelFiles();
            } else {
                convertExcelToWord();
            }
            showAlert("Готово", "Файл успешно создан!");
        } catch (IOException e) {
            showAlert("Ошибка", "Ошибка при обработке: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void mergeExcelFiles() throws IOException {
        String[] filePaths = fileList.toString().split("\n");

        // Создаём новый Workbook для объединения
        try (Workbook mergedWorkbook = new XSSFWorkbook(); FileOutputStream fos = new FileOutputStream(outputFile)) {

            // Создаём новый лист для данных
            Sheet mergedSheet = mergedWorkbook.createSheet("Объединённые данные");

            int rowIndex = 0; // Индекс для строк в mergedSheet
            boolean isHeaderCopied = false; // Флаг для отслеживания, был ли заголовок уже скопирован

            for (String filePath : filePaths) {
                File excelFile = new File(filePath);
                try (FileInputStream fis = new FileInputStream(excelFile); Workbook workbook = new XSSFWorkbook(fis)) {

                    // Пройдем по всем листам текущего файла
                    for (Sheet sheet : workbook) {

                        // Если заголовок еще не скопирован, копируем первую строку
                        if (!isHeaderCopied) {
                            Row headerRow = sheet.getRow(0); // Первая строка - это заголовок
                            if (headerRow != null) {
                                Row newHeaderRow = mergedSheet.createRow(rowIndex++);
                                copyRow(headerRow, newHeaderRow); // Копируем заголовок
                            }
                            isHeaderCopied = true; // Заголовок скопирован
                        }

                        // Пройдем по всем строкам в текущем листе, начиная со второй строки
                        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                            Row row = sheet.getRow(i);
                            Row newRow = mergedSheet.createRow(rowIndex++);
                            copyRow(row, newRow); // Копируем данные строки
                        }
                    }
                }
            }

            // Записываем объединённую книгу в файл
            mergedWorkbook.write(fos);
        }
    }

    private void copyRow(Row sourceRow, Row targetRow) {
        // Копируем данные из каждой ячейки
        for (int i = 0; i < sourceRow.getPhysicalNumberOfCells(); i++) {
            org.apache.poi.ss.usermodel.Cell sourceCell = sourceRow.getCell(i);
            org.apache.poi.ss.usermodel.Cell targetCell = targetRow.createCell(i);

            // В зависимости от типа ячейки копируем данные
            switch (sourceCell.getCellType()) {
                case STRING:
                    targetCell.setCellValue(sourceCell.getStringCellValue());
                    break;
                case NUMERIC:
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    targetCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    targetCell.setCellFormula(sourceCell.getCellFormula());
                    break;
                case BLANK:
                    // Если ячейка пустая, не нужно ничего делать
                    break;
                default:
                    throw new IllegalArgumentException("Неизвестный тип ячейки: " + sourceCell.getCellType());
            }
        }
    }


    private void convertExcelToWord() throws IOException {
        String[] filePaths = fileList.toString().split("\n");

        try (XWPFDocument document = new XWPFDocument(); FileOutputStream fos = new FileOutputStream(outputFile)) {
            boolean isHeaderWritten = false; // Флаг для отслеживания, был ли заголовок записан

            for (String filePath : filePaths) {
                File excelFile = new File(filePath);
                try (FileInputStream fis = new FileInputStream(excelFile); Workbook workbook = new XSSFWorkbook(fis)) {
                    for (Sheet sheet : workbook) {
                        if (!isHeaderWritten) {
                            // Добавляем заголовок только один раз для первого листа
                            document.createParagraph().createRun().setText("Таблица: " + sheet.getSheetName());
                            isHeaderWritten = true;
                        }

                        XWPFTable table = document.createTable();

                        boolean isFirstRow = true; // Флаг для отслеживания первой строки (заголовка)

                        for (Row row : sheet) {
                            XWPFTableRow tableRow = table.createRow();

                            // Если это первая строка, то это заголовок таблицы
                            if (isFirstRow) {
                                isFirstRow = false;
                                // Пропускаем первую строку, не добавляем пустые ячейки
                                continue;
                            }

                            // Копируем данные из всех ячеек в строках
                            for (org.apache.poi.ss.usermodel.Cell cell : row) {
                                XWPFTableCell tableCell = tableRow.createCell(); // Создаем ячейку
                                tableCell.setText(cell.toString().trim()); // Убираем лишние пробелы
                            }
                        }
                    }
                }
            }
            document.write(fos);
        }
    }





    @FXML
    protected void onOpenFile() {
        if (outputFile != null && outputFile.exists()) {
            try {
                // Открываем папку, содержащую файл, с помощью стандартного приложения системы
                Desktop.getDesktop().open(outputFile.getParentFile());
            } catch (IOException e) {
                showAlert("Ошибка", "Не удалось открыть папку: " + e.getMessage());
            }
        } else {
            showAlert("Ошибка", "Файл не существует.");
        }
    }


    private void showAlert(String title, String message) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }
}