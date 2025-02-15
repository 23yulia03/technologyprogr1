package com.example.technologyprogr1;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.Desktop;
import java.util.List;
import java.util.Optional;

public class HelloController {
    @FXML
    private Label fileListLabel;
    @FXML
    private Label folderPathLabel;

    private StringBuilder fileList = new StringBuilder();
    private File saveFolder;
    private Stage stage;

    public void setStage(Stage stage) {
        this.stage = stage;
    }

    @FXML
    protected void onSelectFiles() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Word & Excel Files", "*.docx", "*.xlsx"));
        List<File> files = fileChooser.showOpenMultipleDialog(stage);
        if (files != null) {
            fileList.setLength(0);
            for (File file : files) {
                fileList.append(file.getAbsolutePath()).append("\n");
            }
            fileListLabel.setText("Выбранные файлы:\n" + fileList.toString());
        }
    }

    private void analyzeFileStructure(List<File> files) {
        int columnCount = 0;
        boolean hasEmptyCells = false;

        for (File file : files) {
            if (file.getName().endsWith(".xlsx")) {
                try (FileInputStream fis = new FileInputStream(file);
                     Workbook workbook = new XSSFWorkbook(fis)) {

                    Sheet sheet = workbook.getSheetAt(0);
                    for (Row row : sheet) {
                        columnCount = Math.max(columnCount, row.getLastCellNum());
                        for (Cell cell : row) {
                            if (false) {
                                hasEmptyCells = true;
                            }
                        }
                    }
                } catch (IOException e) {
                    showAlert("Ошибка анализа", "Не удалось проанализировать файл " + file.getName());
                }
            }
        }

        showAlert("Анализ завершен", "Обнаружено " + columnCount + " столбцов. Пустые ячейки: " + (hasEmptyCells ? "Да" : "Нет"));
    }

    @FXML
    protected void onSelectFolder() {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        saveFolder = directoryChooser.showDialog(stage);
        if (saveFolder != null) {
            folderPathLabel.setText("Папка для сохранения: " + saveFolder.getAbsolutePath());
        }
    }

    @FXML
    protected void onConvert() {
        if (fileList.length() == 0 || saveFolder == null) {
            showAlert("Ошибка", "Выберите файлы и папку для сохранения");
            return;
        }

        // Запрос направления конвертации
        List<String> choices = List.of("Word → Excel", "Excel → Word");
        ChoiceDialog<String> dialog = new ChoiceDialog<>(choices.get(0), choices);
        dialog.setTitle("Выбор формата");
        dialog.setHeaderText("Выберите направление конвертации:");
        dialog.setContentText("Конвертировать:");

        Optional<String> result = dialog.showAndWait();
        if (result.isEmpty()) {
            return; // Пользователь отменил выбор
        }

        String conversionType = result.get();
        try {
            if ("Word → Excel".equals(conversionType)) {
                convertWordToExcel();
            } else {
                convertExcelToWord();
            }
            showAlert("Готово", "Конвертация завершена!");
        } catch (IOException e) {
            showAlert("Ошибка", "Ошибка при конвертации: " + e.getMessage());
            e.printStackTrace();
        }
    }

    @FXML
    protected void onOpenFile() {
        if (saveFolder != null && Desktop.isDesktopSupported()) {
            try {
                Desktop.getDesktop().open(saveFolder);
            } catch (IOException e) {
                showAlert("Ошибка", "Не удалось открыть папку");
            }
        } else {
            showAlert("Ошибка", "Папка не выбрана или не поддерживается");
        }
    }

    private void showAlert(String title, String message) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }

    private void convertWordToExcel() throws IOException {
        String[] filePaths = fileList.toString().split("\n");
        for (String filePath : filePaths) {
            File wordFile = new File(filePath);
            if (!wordFile.getName().endsWith(".docx")) continue;

            File excelFile = new File(saveFolder, wordFile.getName().replace(".docx", ".xlsx"));

            try (FileInputStream fis = new FileInputStream(wordFile);
                 XWPFDocument document = new XWPFDocument(fis);
                 Workbook workbook = new XSSFWorkbook()) {

                Sheet sheet = workbook.createSheet("Word Content");
                int rowIndex = 0;

                for (XWPFParagraph paragraph : document.getParagraphs()) {
                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(0).setCellValue(paragraph.getText());
                }

                try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                    workbook.write(fos);
                }
            }
        }
    }

    private void convertExcelToWord() throws IOException {
        String[] filePaths = fileList.toString().split("\n");
        File wordFile = new File(saveFolder, "Результат.docx");

        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream fos = new FileOutputStream(wordFile)) {
            for (String filePath : filePaths) {
                File excelFile = new File(filePath);
                if (!excelFile.getName().endsWith(".xlsx")) continue;

                try (FileInputStream fis = new FileInputStream(excelFile);
                     Workbook workbook = new XSSFWorkbook(fis)) {

                    for (Sheet sheet : workbook) {
                        for (Row row : sheet) {
                            XWPFParagraph paragraph = document.createParagraph();
                            XWPFRun run = paragraph.createRun();
                            for (Cell cell : row) {
                                run.setText(cell.toString() + " ");
                            }
                        }
                    }
                }
            }
            document.write(fos);
        }
    }
}