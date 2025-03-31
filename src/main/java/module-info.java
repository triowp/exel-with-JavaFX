module com.example.exel {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.ooxml;


    opens com.example.exel to javafx.fxml;
    exports com.example.exel;
}