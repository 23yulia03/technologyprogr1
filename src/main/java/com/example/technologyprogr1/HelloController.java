package com.example.technologyprogr1;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
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
    private File outputFile; // Файл для сохранения результата
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

    @FXML
    protected void onSelectFolder() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Выберите место сохранения и имя файла");

        // Устанавливаем начальное имя файла
        if (fileList.length() > 0) {
            String firstFileName = fileList.toString().split("\n")[0];
            if (firstFileName.endsWith(".docx")) {
                fileChooser.setInitialFileName(firstFileName.replace(".docx", ".xlsx")); // Для Word → Excel
            } else if (firstFileName.endsWith(".xlsx")) {
                fileChooser.setInitialFileName("Результат.docx"); // Для Excel → Word
            }
        }

        // Фильтры для выбора расширения файла
        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Word Documents (*.docx)", "*.docx"),
                new FileChooser.ExtensionFilter("Excel Files (*.xlsx)", "*.xlsx")
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
        if (outputFile != null && Desktop.isDesktopSupported()) {
            try {
                Desktop.getDesktop().open(outputFile.getParentFile()); // Открываем папку с результатом
            } catch (IOException e) {
                showAlert("Ошибка", "Не удалось открыть папку");
            }
        } else {
            showAlert("Ошибка", "Файл не выбран или не поддерживается");
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

            // Используем outputFile для сохранения
            File excelFile = new File(outputFile.getParent(), outputFile.getName());

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

        // Используем outputFile для сохранения
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream fos = new FileOutputStream(outputFile)) {
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