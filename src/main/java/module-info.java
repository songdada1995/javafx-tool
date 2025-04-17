module org.example.javafxtool {
    requires javafx.controls;
    requires javafx.fxml;
    requires commons.lang3;
    requires static lombok;
    requires sevenzipjbinding;
    requires poi;
    requires poi.ooxml;
    requires junrar;
    requires cn.hutool;


    opens org.example.javafxtool to javafx.fxml;
    exports org.example.javafxtool.excel;
    exports org.example.javafxtool.handler;
    exports org.example.javafxtool;
}