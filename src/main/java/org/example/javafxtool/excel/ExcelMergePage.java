package org.example.javafxtool.excel;

import javafx.scene.layout.StackPane;

public class ExcelMergePage {

    private StackPane page;

    public StackPane getPage() {
        if (page == null) {
            page = new StackPane();
        }
        return page;
    }
}