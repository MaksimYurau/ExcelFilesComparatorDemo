module com.example.excelfilescomparatordemo {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;


    opens com.example.excelfilescomparatordemo to javafx.fxml;
    exports com.example.excelfilescomparatordemo;
}