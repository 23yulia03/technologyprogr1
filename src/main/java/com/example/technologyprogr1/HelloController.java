package com.example.technologyprogr1;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.Label;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

public class HelloController {

    @FXML
    private Label fileListLabel;
    @FXML
    private Label folderPathLabel;
    @FXML
    private Label statusLabel;
    private List<File> selectedFiles = new ArrayList<>();
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
            selectedFiles.clear();
            selectedFiles.addAll(files);
            updateFileListLabel();
        }
    }

    @FXML
    private void handleSelectFilesButtonAction() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
        List<File> files = fileChooser.showOpenMultipleDialog(stage);
        if (files != null) {
            selectedFiles.clear();
            selectedFiles.addAll(files);
            updateFileListLabel();
        }
    }

    @FXML
    private void handleSelectFolderButtonAction() {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        File selectedDirectory = directoryChooser.showDialog(stage);
        if (selectedDirectory != null) {
            selectedFiles.clear();
            selectedFiles.addAll(Arrays.asList(selectedDirectory.listFiles((dir, name) -> name.endsWith(".xlsx"))));
            updateFileListLabel();
        }
    }

    private void updateFileListLabel() {
        if (selectedFiles.isEmpty()) {
            fileListLabel.setText("Выбранные файлы: Нет файлов");
        } else {
            StringBuilder paths = new StringBuilder("Выбранные файлы:\n");
            for (File file : selectedFiles) {
                paths.append(file.getAbsolutePath()).append("\n");
            }
            fileListLabel.setText(paths.toString());
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
            List<String> data = ExcelReader.readExcelFiles(selectedFiles);
            if ("Excel → Excel".equals(result.get())) {
                ExcelWriter.writeToExcel(data, outputFile);
            } else {
                WordWriter.writeToWord(data, outputFile);
            }
            showAlert("Готово", "Файл успешно создан!");
        } catch (IOException e) {
            showAlert("Ошибка", "Ошибка при обработке: " + e.getMessage());
            e.printStackTrace();
        }
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

    private void showAlert(String title, String message) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }
}

