package org.example.javafxtool.excel;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Menu;
import javafx.scene.control.MenuBar;
import javafx.scene.control.RadioMenuItem;
import javafx.scene.control.ToggleGroup;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;

import java.net.URL;

public class ExcelToolkitApp extends Application {

    private BorderPane mainPane;
    private ToggleGroup menuToggleGroup;
    private StackPane renamePage;
    private StackPane mergePage;

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Excel 工具集");

        // 初始化菜单栏
        menuToggleGroup = new ToggleGroup();
        RadioMenuItem renameMenuItem = new RadioMenuItem("Excel 重命名");
        renameMenuItem.setToggleGroup(menuToggleGroup);
        renameMenuItem.setSelected(true);
        RadioMenuItem mergeMenuItem = new RadioMenuItem("Excel 合并");
        mergeMenuItem.setToggleGroup(menuToggleGroup);

        Menu menu = new Menu("菜单");
        menu.getItems().addAll(renameMenuItem, mergeMenuItem);
        MenuBar menuBar = new MenuBar();
        menuBar.getMenus().add(menu);

        // 初始化页面
        renamePage = new ExcelRenamePage().getPage();
        mergePage = new ExcelMergePage().getPage();

        // 切换页面逻辑
        renameMenuItem.setOnAction(e -> showPage(renamePage));
        mergeMenuItem.setOnAction(e -> showPage(mergePage));

        mainPane = new BorderPane();
        mainPane.setTop(menuBar);
        mainPane.setCenter(renamePage);

        Scene scene = new Scene(mainPane, 800, 600);
        URL resource = getClass().getResource("/styles/main.css");
        if (resource != null) {
            scene.getStylesheets().add(resource.toExternalForm());
        } else {
            System.out.println("未找到 main.css 文件，请检查文件路径。");
        }
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void showPage(StackPane page) {
        mainPane.setCenter(page);
    }

    public static void main(String[] args) {
        launch(args);
    }
}