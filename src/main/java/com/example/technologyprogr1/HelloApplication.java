package com.example.technologyprogr1;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class HelloApplication extends Application {
    @Override
    public void start(Stage primaryStage) throws Exception {
        FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("hello-view.fxml"));
        Scene scene = new Scene(fxmlLoader.load(), 400, 250);
        primaryStage.setTitle("Конвертер Word ↔️ Excel");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}