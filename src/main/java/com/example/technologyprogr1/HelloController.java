package com.example.technologyprogr1;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.Label;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.awt.*;
import java.awt.TextArea;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

public class HelloController {
    @FXML
    private Label fileListLabel;
    @FXML
    private Label folderPathLabel;
    @FXML
    private TextArea textArea; // Для вывода текста в GUI

    private List<File> selectedFiles = new ArrayList<>();
    private File outputFile;
    private Stage stage;

    // Ожидаемые заголовки и количество столбцов
    private final List<String> expectedHeaders = List.of("Номер", "Цена", "Количество"); // Замените на реальные заголовки
    private final int expectedColumns = expectedHeaders.size();

    public void setStage(Stage stage) {
        this.stage = stage;
    }

    @FXML
    protected void onSelectFiles() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
        List<File> files = fileChooser.showOpenMultipleDialog(stage);
        if (files != null) {
            selectedFiles.clear();
            selectedFiles.addAll(files);
            fileListLabel.setText("Выбранные файлы:\n" + files.size());
        }
        if (files != null) {
            StringBuilder fileList = new StringBuilder(); // Создаем объект StringBuilder
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
        fileChooser.setInitialFileName("Результат");
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
        if (selectedFiles.isEmpty() || outputFile == null) {
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
        try (Workbook mergedWorkbook = new XSSFWorkbook(); FileOutputStream fos = new FileOutputStream(outputFile)) {
            Sheet mergedSheet = mergedWorkbook.createSheet("Объединённые данные");
            int rowIndex = 0;
            boolean isHeaderCopied = false;

            for (File file : selectedFiles) {
                try (FileInputStream fis = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(fis)) {
                    for (Sheet sheet : workbook) {
                        if (!isHeaderCopied) {
                            Row headerRow = sheet.getRow(0);
                            if (headerRow != null) {
                                Row newHeaderRow = mergedSheet.createRow(rowIndex++);
                                copyRow(headerRow, newHeaderRow);
                            }
                            isHeaderCopied = true;
                        }
                        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                            Row row = sheet.getRow(i);
                            if (row != null) {
                                Row newRow = mergedSheet.createRow(rowIndex++);
                                copyRow(row, newRow);
                            }
                        }
                    }
                } catch (IOException e) {
                    showAlert("Ошибка", "Ошибка при чтении файла: " + file.getName());
                    throw e;
                }
            }
            mergedWorkbook.write(fos);
        } catch (IOException e) {
            showAlert("Ошибка", "Ошибка при создании файла: " + e.getMessage());
            throw e;
        }
    }

    private void copyRow(Row sourceRow, Row targetRow) {
        for (int i = 0; i < sourceRow.getPhysicalNumberOfCells(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell targetCell = targetRow.createCell(i);
            switch (sourceCell.getCellType()) {
                case STRING -> targetCell.setCellValue(sourceCell.getStringCellValue());
                case NUMERIC -> targetCell.setCellValue(sourceCell.getNumericCellValue());
                case BOOLEAN -> targetCell.setCellValue(sourceCell.getBooleanCellValue());
                case FORMULA -> targetCell.setCellFormula(sourceCell.getCellFormula());
            }
        }
    }

    private void convertExcelToWord() throws IOException {
        try (XWPFDocument document = new XWPFDocument()) {
            XWPFTable table = document.createTable(); // Создаем таблицу сразу
            if (table.getRows().size() > 0) {
                table.removeRow(0); // Удаляем дефолтную строку, если она есть
            }
            createHeaderRow(table, expectedHeaders); // Заполняем заголовки

            boolean isFirstFile = true;

            for (File file : selectedFiles) {
                try (FileInputStream fis = new FileInputStream(file);
                     Workbook workbook = new XSSFWorkbook(fis)) {

                    Sheet sheet = workbook.getSheetAt(0);
                    if (!validateColumns(sheet, expectedColumns, expectedHeaders)) {
                        showAlert("Ошибка", "Недопустимый заголовок в: " + file.getName());
                        continue;
                    }

                    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);
                        if (row == null) continue;

                        if (!isFirstFile && i == 0) continue; // Пропускаем заголовки после первого файла

                        XWPFTableRow tableRow = table.createRow();
                        for (int j = 0; j < expectedColumns; j++) {
                            Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            String value = cellToString(cell);

                            // Если ячейка не существует, создаем ее
                            if (tableRow.getCell(j) == null) {
                                tableRow.addNewTableCell();
                            }
                            tableRow.getCell(j).setText(value);
                        }
                    }
                    isFirstFile = false;
                } catch (Exception e) {
                    e.printStackTrace();
                    showAlert("Ошибка", "Ошибка при чтении файла: " + file.getName() + " - " + e.getMessage());
                }
            }

            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                document.write(out);
            }
        } catch (Exception e) {
            e.printStackTrace();
            showAlert("Ошибка", "Ошибка при конвертации в Word: " + e.getMessage());
        }
    }

    private void createHeaderRow(XWPFTable table, List<String> headers) {
        XWPFTableRow headerRow = table.createRow();
        for (int i = 0; i < headers.size(); i++) {
            if (headerRow.getCell(i) == null) {
                headerRow.addNewTableCell();
            }
            headerRow.getCell(i).setText(headers.get(i));
        }
    }

    private String cellToString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double value = cell.getNumericCellValue();
                    if (value == (long) value) {
                        return String.format("%d", (long) value);
                    } else {
                        return String.valueOf(value);
                    }
                }
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return "";
        }
    }

    private boolean validateColumns(Sheet sheet, int expectedColumns, List<String> expectedHeaders) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null || headerRow.getLastCellNum() < expectedColumns) return false;

        for (int i = 0; i < expectedHeaders.size(); i++) {
            Cell cell = headerRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            String actualHeader = cellToString(cell).trim();
            String expectedHeader = expectedHeaders.get(i).trim();
            if (!actualHeader.equals(expectedHeader)) {
                return false;
            }
        }
        return true;
    }

    @FXML
    protected void onOpenFile() {
        if (outputFile != null && outputFile.exists()) {
            try {
                Desktop.getDesktop().open(outputFile.getParentFile());
            } catch (IOException e) {
                showAlert("Ошибка", "Не удалось открыть папку: " + e.getMessage());
            }
        } else {
            showAlert("Ошибка", "Файл не существует.");
        }
    }

    @FXML
    protected void handleViewFile() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Открыть файл для просмотра");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Word Files", "*.docx"));
        File file = fileChooser.showOpenDialog(null);
        if (file != null) {
            try (XWPFDocument document = new XWPFDocument(new FileInputStream(file))) {
                StringBuilder content = new StringBuilder();
                for (XWPFParagraph paragraph : document.getParagraphs()) {
                    content.append(paragraph.getText()).append("\n");
                }
                textArea.setText(content.toString());
            } catch (IOException e) {
                textArea.appendText("Ошибка при чтении файла: " + e.getMessage() + "\n");
            }
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