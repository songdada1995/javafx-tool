package org.example.javafxtool.handler;

import javafx.animation.PauseTransition;
import javafx.application.Application;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Duration;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.github.junrar.Archive;
import com.github.junrar.exception.RarException;
import com.github.junrar.rarfile.FileHeader;

import java.io.*;
import java.net.URL;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class ExcelRenameTool extends Application {

    private TextField rowInput;
    private TextField colInput;
    private TextField sheetNameInput;
    private Button fileSelectButton;
    private Label fileNameLabel;
    private Label rowErrorLabel;
    private Label colErrorLabel;
    private Label sheetNameErrorLabel;
    private ProgressIndicator loadingIndicator;
    private Label resultLabel;
    private Label fileSelectErrorLabel;
    private File selectedFile;
    private StackPane dragDropPane;
    private Label dragDropHintLabel;
    private Label descriptionLabel;
    private StackPane overlayPane;
    private StackPane root;

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Excel 重命名工具");

        // 优化字体和排版的描述信息
        descriptionLabel = new Label("提取单元格数据作为文件名，并重命名文件");
        descriptionLabel.setStyle("-fx-font-family: 'Open Sans', sans-serif; -fx-font-size: 18px; -fx-font-weight: bold; -fx-padding: 3px 0 25px 0; -fx-alignment: center;");

        // 创建输入框和标签，添加星号标识必填
        Label rowLabel = new Label("* 行号");
        rowInput = new TextField();
        rowInput.setPromptText("请输入正整数");
        rowErrorLabel = new Label();
        rowErrorLabel.setTextFill(Color.RED);
        rowErrorLabel.setOpacity(0);

        Label colLabel = new Label("* 列号");
        colInput = new TextField();
        colInput.setPromptText("请输入英文字母");
        colErrorLabel = new Label();
        colErrorLabel.setTextFill(Color.RED);
        colErrorLabel.setOpacity(0);

        Label sheetNameLabel = new Label("sheet 名称");
        sheetNameInput = new TextField();
        sheetNameInput.setPromptText("可选");
        sheetNameErrorLabel = new Label();
        sheetNameErrorLabel.setTextFill(Color.RED);
        sheetNameErrorLabel.setOpacity(0);

        // 输入框失焦校验
        rowInput.focusedProperty().addListener((observable, oldValue, newValue) -> {
            if (!newValue) {
                validateRowInput();
            }
        });
        colInput.focusedProperty().addListener((observable, oldValue, newValue) -> {
            if (!newValue) {
                validateColInput();
            }
        });

        // 创建文件选择按钮，添加提示语
        Label fileSelectStarLabel = new Label("*");
        fileSelectButton = new Button("选择文件");
        fileSelectButton.getStyleClass().add("file-select-button-style");
        Tooltip fileSelectTooltip = new Tooltip("请选择 rar/zip 文件");
        fileSelectButton.setTooltip(fileSelectTooltip);
        fileSelectButton.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            fileChooser.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("压缩文件", "*.zip", "*.rar")
            );
            selectedFile = fileChooser.showOpenDialog(primaryStage);
            if (selectedFile != null) {
                fileNameLabel.setText(selectedFile.getName());
                dragDropHintLabel.setVisible(false);
                fileSelectErrorLabel.setOpacity(0);
                setDragDropPaneBorderColor(Color.GREEN);
            }
        });

        // 创建拖拽虚线框
        dragDropPane = new StackPane();
        dragDropPane.setBorder(new Border(new BorderStroke(Color.GRAY, BorderStrokeStyle.DASHED, new CornerRadii(5), new BorderWidths(1))));
        dragDropPane.setPadding(new Insets(10));
        dragDropPane.setAlignment(Pos.CENTER);
        dragDropPane.setMinWidth(250);
        fileNameLabel = new Label();
        fileNameLabel.setMaxWidth(250);
        fileNameLabel.setEllipsisString("...");
        Tooltip tooltip = new Tooltip();
        fileNameLabel.setTooltip(tooltip);
        fileNameLabel.textProperty().addListener((observable, oldValue, newValue) -> {
            tooltip.setText(newValue);
        });
        dragDropHintLabel = new Label("拖拽文件到此处");
        dragDropPane.getChildren().addAll(dragDropHintLabel, fileNameLabel);

        // 支持文件拖拽到虚线框
        dragDropPane.setOnDragOver(this::handleDragOver);
        dragDropPane.setOnDragDropped(this::handleDragDropped);

        // 创建确定按钮
        Button confirmButton = new Button("确定");
        confirmButton.getStyleClass().add("file-select-button-style");
        confirmButton.setOnAction(e -> {
            if (validateInputs()) {
                showLoading();
                new Thread(() -> {
                    boolean result = processFile();
                    Platform.runLater(() -> {
                        hideLoading();
                        // 延迟一小段时间再显示结果，确保loading框和蒙版已隐藏
                        PauseTransition delay = new PauseTransition(Duration.millis(100));
                        delay.setOnFinished(event -> {
                            showResult(result);
                        });
                        delay.play();
                    });
                }).start();
            }
        });

        // 创建重置按钮
        Button resetButton = new Button("重置");
        resetButton.getStyleClass().add("file-select-button-style");
        resetButton.setOnAction(e -> {
            rowInput.clear();
            colInput.clear();
            sheetNameInput.clear();
            selectedFile = null;
            fileNameLabel.setText("");
            dragDropHintLabel.setVisible(true);
            rowErrorLabel.setOpacity(0);
            colErrorLabel.setOpacity(0);
            sheetNameErrorLabel.setOpacity(0);
            resultLabel.setText("");
            fileSelectErrorLabel.setOpacity(0);
            setDragDropPaneBorderColor(Color.GRAY);
            hideResult();
        });

        resultLabel = new Label();
        loadingIndicator = new ProgressIndicator();
        loadingIndicator.setVisible(false);
        fileSelectErrorLabel = new Label();
        fileSelectErrorLabel.setTextFill(Color.RED);
        fileSelectErrorLabel.setOpacity(0);

        // 布局设置
        GridPane grid = new GridPane();
        grid.setAlignment(Pos.CENTER);
        grid.setHgap(10);
        grid.setVgap(3);
        grid.setPadding(new Insets(30, 25, 30, 25));
        grid.getStyleClass().add("grid-pane-style");

        grid.add(descriptionLabel, 0, 0, 2, 1);
        grid.add(rowLabel, 0, 1);
        grid.add(rowInput, 1, 1);
        grid.add(rowErrorLabel, 1, 2);
        grid.add(colLabel, 0, 3);
        grid.add(colInput, 1, 3);
        grid.add(colErrorLabel, 1, 4);
        grid.add(fileSelectButton, 0, 5);
        grid.add(dragDropPane, 1, 5);
        grid.add(fileSelectErrorLabel, 1, 6);
        grid.add(sheetNameLabel, 0, 7);
        grid.add(sheetNameInput, 1, 7);
        grid.add(sheetNameErrorLabel, 1, 8);
        grid.add(confirmButton, 0, 10);
        grid.add(resetButton, 1, 10);

        // 创建 overlay 面板用于显示 loading 和结果
        overlayPane = new StackPane();
        overlayPane.setBackground(new Background(new BackgroundFill(Color.rgb(0, 0, 0, 0.3), CornerRadii.EMPTY, Insets.EMPTY)));
        overlayPane.getChildren().add(loadingIndicator);
        overlayPane.setVisible(false);

        root = new StackPane();
        root.getChildren().addAll(grid, overlayPane, resultLabel);
        resultLabel.setVisible(false);

        // 点击面板空白处触发输入框校验
        root.setOnMouseClicked(event -> {
            if (!rowInput.isFocused()) {
                validateRowInput();
            }
            if (!colInput.isFocused()) {
                validateColInput();
            }
        });

        Scene scene = new Scene(root, 600, 500);
        URL resource = getClass().getResource("/styles.css");
        if (resource != null) {
            scene.getStylesheets().add(resource.toExternalForm());
        } else {
            System.out.println("未找到 styles.css 文件，请检查文件路径。");
        }
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private boolean validateRowInput() {
        String rowText = rowInput.getText().trim();
        if (rowText.isEmpty()) {
            rowErrorLabel.setText("行号不能为空");
            rowErrorLabel.setOpacity(1);
            return false;
        } else {
            try {
                int row = Integer.parseInt(rowText);
                if (row <= 0) {
                    rowErrorLabel.setText("行号必须为正整数");
                    rowErrorLabel.setOpacity(1);
                    return false;
                } else {
                    rowErrorLabel.setOpacity(0);
                    return true;
                }
            } catch (NumberFormatException e) {
                rowErrorLabel.setText("行号必须为正整数");
                rowErrorLabel.setOpacity(1);
                return false;
            }
        }
    }

    private boolean validateColInput() {
        String colText = colInput.getText().trim();
        if (colText.isEmpty()) {
            colErrorLabel.setText("列号不能为空");
            colErrorLabel.setOpacity(1);
            return false;
        } else if (!colText.matches("[a-zA-Z]+")) {
            colErrorLabel.setText("列号必须为英文字母");
            colErrorLabel.setOpacity(1);
            return false;
        } else {
            colInput.setText(colText.toUpperCase());
            colErrorLabel.setOpacity(0);
            return true;
        }
    }

    private boolean validateInputs() {
        boolean isValidRow = validateRowInput();
        boolean isValidCol = validateColInput();
        if (selectedFile == null) {
            fileSelectErrorLabel.setText("请选择文件");
            fileSelectErrorLabel.setOpacity(1);
            return false;
        } else {
            fileSelectErrorLabel.setOpacity(0);
        }
        return isValidRow && isValidCol;
    }

    private boolean processFile() {
        if (selectedFile == null) {
            return false;
        }
        try {
            int rowIndex = Integer.parseInt(rowInput.getText().trim()) - 1;
            int colIndex = convertColumnToIndex(colInput.getText().trim());
            String sheetName = sheetNameInput.getText().trim();

            File outputDir = new File(selectedFile.getParent(), "重命名文件");
            if (!outputDir.exists()) {
                outputDir.mkdirs();
            }

            Thread.sleep(3 * 1000);

            if (selectedFile.getName().endsWith(".zip")) {
                try (ZipFile zipFile = new ZipFile(selectedFile)) {
                    Enumeration<? extends ZipEntry> entries = zipFile.entries();
                    while (entries.hasMoreElements()) {
                        ZipEntry entry = entries.nextElement();
                        if (entry.getName().endsWith(".xlsx")) {
                            try (InputStream inputStream = zipFile.getInputStream(entry)) {
                                processExcel(inputStream, outputDir, rowIndex, colIndex, sheetName);
                            }
                        }
                    }
                }
            } else if (selectedFile.getName().endsWith(".rar")) {
                try (Archive archive = new Archive(selectedFile)) {
                    FileHeader fileHeader;
                    while ((fileHeader = archive.nextFileHeader()) != null) {
                        if (fileHeader.getFileNameW().endsWith(".xlsx")) {
                            try (InputStream inputStream = archive.getInputStream(fileHeader)) {
                                processExcel(inputStream, outputDir, rowIndex, colIndex, sheetName);
                            }
                        }
                    }
                }
            }
            return true;
        } catch (IOException | RarException e) {
            e.printStackTrace();
            return false;
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
    }

    private void processExcel(InputStream inputStream, File outputDir, int rowIndex, int colIndex, String sheetName) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet;
            if (sheetName.isEmpty()) {
                sheet = workbook.getSheetAt(0);
            } else {
                sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    return;
                }
            }
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(colIndex);
                if (cell != null) {
                    String newFileName = cell.getStringCellValue() + ".xlsx";
                    File outputFile = new File(outputDir, newFileName);
                    try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
                        workbook.write(outputStream);
                    }
                }
            }
        }
    }

    private int convertColumnToIndex(String column) {
        int index = 0;
        for (int i = 0; i < column.length(); i++) {
            index = index * 26 + (column.charAt(i) - 'A' + 1);
        }
        return index - 1;
    }

    private void showLoading() {
        overlayPane.setVisible(true);
        loadingIndicator.setVisible(true);
    }

    private void hideLoading() {
        overlayPane.setVisible(false);
        loadingIndicator.setVisible(false);
    }

    private void showResult(boolean success) {
        resultLabel.setVisible(true);
        resultLabel.setText(success ? "处理成功" : "处理失败");
        resultLabel.setTextFill(success ? Color.GREEN : Color.RED);
        // 设置字体和行号字体大小一致，使用和说明文字相同的字体
        resultLabel.setStyle("-fx-font-family: 'Open Sans', sans-serif; -fx-font-size: 18px; -fx-padding: 20px; -fx-background-color: white; -fx-border-radius: 5px; -fx-effect: dropshadow(gaussian, rgba(0, 0, 0, 0.2), 10, 0, 0, 3); -fx-alignment: center;");

        PauseTransition pause = new PauseTransition(Duration.seconds(2));
        pause.setOnFinished(e -> {
            resultLabel.setVisible(false);
        });
        pause.play();
    }

    private void hideResult() {
        resultLabel.setVisible(false);
    }

    private void handleDragOver(DragEvent event) {
        if (event.getDragboard().hasFiles()) {
            event.acceptTransferModes(TransferMode.COPY);
        }
        event.consume();
    }

    private void handleDragDropped(DragEvent event) {
        Dragboard db = event.getDragboard();
        boolean success = false;
        if (db.hasFiles()) {
            selectedFile = db.getFiles().get(0);
            if (selectedFile.getName().endsWith(".zip") || selectedFile.getName().endsWith(".rar")) {
                fileNameLabel.setText(selectedFile.getName());
                dragDropHintLabel.setVisible(false);
                fileSelectErrorLabel.setOpacity(0);
                setDragDropPaneBorderColor(Color.GREEN);
                success = true;
            }
        }
        event.setDropCompleted(success);
        event.consume();
    }

    private void setDragDropPaneBorderColor(Color color) {
        dragDropPane.setBorder(new Border(new BorderStroke(color, BorderStrokeStyle.DASHED, new CornerRadii(5), new BorderWidths(1))));
    }

    public static void main(String[] args) {
        launch(args);
    }
}