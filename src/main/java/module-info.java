module com.example.technologyprogr1 {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.ooxml;
    requires java.desktop;


    opens com.example.technologyprogr1 to javafx.fxml;
    exports com.example.technologyprogr1;
}