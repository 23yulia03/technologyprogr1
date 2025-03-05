package com.example.technologyprogr1;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

public class HelloController {

    @FXML
    private Label fileListLabel; // Метка для отображения выбранных файлов
    @FXML
    private Label folderPathLabel; // Метка для отображения пути сохранения
    @FXML
    private Label formatLabel; // Метка для отображения выбранного формата

    private List<File> selectedFiles = new ArrayList<>(); // Список выбранных файлов
    private File outputFile; // Файл для сохранения результата
    private Stage stage; // Окно приложения
    private String selectedFormat = ""; // Выбранный формат ("Excel → Excel" или "Excel → Word")

    // Устанавливаем окно приложения
    public void setStage(Stage stage) {
        this.stage = stage;
    }

    @FXML
    public void initialize() {
        fileListLabel.setText("Выбранные файлы: Нет файлов");
    }

    // Выбор формата итогового файла
    @FXML
    protected void onSelectFormat() {
        List<String> choices = List.of("Excel → Excel", "Excel → Word");
        ChoiceDialog<String> dialog = new ChoiceDialog<>(choices.get(0), choices);
        dialog.setTitle("Выбор формата");
        dialog.setHeaderText("Выберите формат итогового файла:");
        dialog.setContentText("Формат:");

        Optional<String> result = dialog.showAndWait();
        if (result.isPresent()) {
            selectedFormat = result.get();
            formatLabel.setText("Выбранный формат: " + selectedFormat); // Обновляем метку
            selectedFiles.clear(); // Очищаем список выбранных файлов
            fileListLabel.setText("Выбранные файлы: Нет файлов"); // Обновляем метку
            folderPathLabel.setText("Папка для сохранения: Не выбрана"); // Обновляем метку
        }
    }

    // Выбор файлов Excel
    @FXML
    protected void onSelectFiles() {
        if (selectedFormat.isEmpty()) {
            showAlert("Ошибка", "Сначала выберите формат итогового файла.");
            return;
        }

        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx")); // Фильтр для Excel
        List<File> files = fileChooser.showOpenMultipleDialog(stage);
        if (files != null) {
            for (File file : files) {
                if (!file.getName().endsWith(".xlsx")) {
                    showAlert("Ошибка", "Выбранный файл " + file.getName() + " не является файлом Excel.");
                    return;
                }
            }
            selectedFiles.clear();
            selectedFiles.addAll(files);
            System.out.println("Выбраны файлы: " + selectedFiles); // Логирование
            updateFileListLabel(); // Обновляем интерфейс
        }
    }

    // Выбор папки с файлами Excel
    @FXML
    protected void onSelectFolder() {
        if (selectedFormat.isEmpty()) {
            showAlert("Ошибка", "Сначала выберите формат итогового файла.");
            return;
        }

        DirectoryChooser directoryChooser = new DirectoryChooser();
        File selectedDirectory = directoryChooser.showDialog(stage);
        if (selectedDirectory != null) {
            File[] files = selectedDirectory.listFiles((dir, name) -> name.endsWith(".xlsx")); // Фильтр для Excel
            if (files == null || files.length == 0) {
                showAlert("Ошибка", "В выбранной папке нет файлов Excel.");
                return;
            }
            selectedFiles.clear();
            selectedFiles.addAll(Arrays.asList(files));
            System.out.println("Выбраны файлы из папки: " + selectedFiles); // Логирование
            updateFileListLabel(); // Обновляем интерфейс
        }
    }

    // Обновление метки с выбранными файлами
    private void updateFileListLabel() {
        if (selectedFiles.isEmpty()) {
            fileListLabel.setText("Выбранные файлы: 0");
        } else {
            StringBuilder paths = new StringBuilder("Выбранные файлы (" + selectedFiles.size() + "):\n");
            for (File file : selectedFiles) {
                paths.append(file.getAbsolutePath()).append("\n");
            }
            fileListLabel.setText(paths.toString());
        }
        System.out.println("Обновлена метка файлов:\n" + fileListLabel.getText()); // Логирование
    }

    // Выбор места сохранения итогового файла
    @FXML
    protected void onSelectOutputFile() {
        if (selectedFormat.isEmpty()) {
            showAlert("Ошибка", "Сначала выберите формат итогового файла.");
            return;
        }

        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Выберите место сохранения и имя файла");
        fileChooser.setInitialFileName("Результат");
        if (selectedFormat.equals("Excel → Excel")) {
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files (*.xlsx)", "*.xlsx"));
        } else {
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Word Documents (*.docx)", "*.docx"));
        }

        outputFile = fileChooser.showSaveDialog(stage);
        if (outputFile != null) {
            folderPathLabel.setText("Файл будет сохранен как: " + outputFile.getAbsolutePath());
        }
    }

    // Запуск конвертации
    @FXML
    protected void onConvert() {
        if (selectedFiles.isEmpty() || outputFile == null) {
            showAlert("Ошибка", "Выберите файлы и укажите имя файла для сохранения.");
            return;
        }

        try {
            List<String> data = ExcelReader.readExcelFiles(selectedFiles);
            if (selectedFormat.equals("Excel → Excel")) {
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

    // Открытие папки с результатом
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

    // Показ сообщения об ошибке
    private void showAlert(String title, String message) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }
}
