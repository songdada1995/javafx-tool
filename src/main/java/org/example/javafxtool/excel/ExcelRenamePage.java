package org.example.javafxtool.excel;

import com.github.junrar.Archive;
import com.github.junrar.exception.RarException;
import com.github.junrar.rarfile.FileHeader;
import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.*;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.stage.FileChooser;
import javafx.util.Duration;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class ExcelRenamePage {

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

    public StackPane getPage() {
        initPage();
        return root;
    }

    private void initPage() {
        // 优化字体和排版的描述信息
        descriptionLabel = new Label("提取单元格数据作为文件名，并重命名文件");
        descriptionLabel.getStyleClass().add("description-label");

        // 创建输入框和标签，添加星号标识必填
        Label rowLabel = new Label("* 行号");
        rowLabel.setAlignment(Pos.CENTER_RIGHT);
        rowInput = new TextField();
        rowInput.setPromptText("请输入正整数");
        rowInput.setAlignment(Pos.CENTER_LEFT);
        rowErrorLabel = new Label();
        rowErrorLabel.getStyleClass().add("error-label");
        rowErrorLabel.setOpacity(0);

        Label colLabel = new Label("* 列号");
        colLabel.setAlignment(Pos.CENTER_RIGHT);
        colInput = new TextField();
        colInput.setPromptText("请输入英文字母");
        colInput.setAlignment(Pos.CENTER_LEFT);
        colErrorLabel = new Label();
        colErrorLabel.getStyleClass().add("error-label");
        colErrorLabel.setOpacity(0);

        Label sheetNameLabel = new Label("sheet 名称");
        sheetNameLabel.setAlignment(Pos.CENTER_RIGHT);
        sheetNameInput = new TextField();
        sheetNameInput.setPromptText("选填，默认提取第1页");
        sheetNameInput.setAlignment(Pos.CENTER_LEFT);
        sheetNameErrorLabel = new Label();
        sheetNameErrorLabel.getStyleClass().add("error-label");
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
        fileSelectButton.getStyleClass().add("select-button-style");
        Tooltip fileSelectTooltip = new Tooltip("请选择 rar/zip 文件");
        fileSelectButton.setTooltip(fileSelectTooltip);
        fileSelectButton.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            fileChooser.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("压缩文件", "*.zip", "*.rar")
            );
            selectedFile = fileChooser.showOpenDialog(null);
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
        confirmButton.getStyleClass().add("button-style");
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
        resetButton.getStyleClass().add("button-style");
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
        resultLabel.getStyleClass().add("result-label");
        loadingIndicator = new ProgressIndicator();
        loadingIndicator.setVisible(false);
        fileSelectErrorLabel = new Label();
        fileSelectErrorLabel.getStyleClass().add("error-label");
        fileSelectErrorLabel.setOpacity(0);

        // 布局设置
        GridPane grid = new GridPane();
        grid.getStyleClass().add("grid-pane-style");
        grid.setAlignment(Pos.CENTER);
        grid.setHgap(4);
        grid.setVgap(9);
        grid.setPadding(new Insets(30, 25, 30, 25));

        grid.add(descriptionLabel, 0, 0, 2, 1);
        grid.add(rowLabel, 0, 1);
        grid.add(rowInput, 1, 1);
        grid.add(rowErrorLabel, 1, 2);
        grid.add(colLabel, 0, 3);
        grid.add(colInput, 1, 3);
        grid.add(colErrorLabel, 1, 4);

//        HBox fileSelectBox = new HBox(5, fileSelectStarLabel, fileSelectButton);
//        grid.add(fileSelectBox, 0, 5);
//        grid.add(dragDropPane, 1, 5);
//        grid.add(fileSelectErrorLabel, 1, 6);

        grid.add(fileSelectButton, 0, 5);
        grid.add(dragDropPane, 1, 5);
        grid.add(fileSelectErrorLabel, 1, 6);
        grid.add(sheetNameLabel, 0, 7);
        grid.add(sheetNameInput, 1, 7);
        grid.add(sheetNameErrorLabel, 1, 8);
        grid.add(confirmButton, 0, 10);
        grid.add(resetButton, 1, 10, GridPane.REMAINING, 1);
        GridPane.setHalignment(resetButton, HPos.RIGHT);

        // 创建 overlay 面板用于显示 loading 和结果
        overlayPane = new StackPane();
        overlayPane.getStyleClass().add("overlay-pane");
        overlayPane.getChildren().add(loadingIndicator);
        overlayPane.setVisible(false);

        root = new StackPane();
        root.getChildren().addAll(grid, overlayPane, resultLabel);
        resultLabel.setVisible(false);
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
}