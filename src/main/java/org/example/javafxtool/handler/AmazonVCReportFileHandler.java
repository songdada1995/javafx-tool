package org.example.javafxtool.handler;

import javafx.application.Application;
import javafx.concurrent.Task;
import javafx.concurrent.WorkerStateEvent;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.commons.lang3.StringUtils;
import org.example.javafxtool.file.FileUtil;

import java.io.File;

public class AmazonVCReportFileHandler extends Application {

    private ProgressBar progressBar;
    private StackPane loadingPane;
    private Label fileLabel;

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("处理Amazon广告发票文件");

        Button selectFileButton = new Button("选择文件");
        fileLabel = new Label("文件: ");

        selectFileButton.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                FileChooser fileChooser = new FileChooser();
                fileChooser.setTitle("打开文件");
                File selectedFile = fileChooser.showOpenDialog(primaryStage);
                if (selectedFile != null) {
                    fileLabel.setText("文件名: " + selectedFile.getName());
                    processFile(selectedFile);
                }
            }
        });

        progressBar = new ProgressBar();
        progressBar.setVisible(false);

        loadingPane = new StackPane();
        loadingPane.setStyle("-fx-background-color: rgba(0, 0, 0, 0.5)");
        loadingPane.setVisible(false);
        loadingPane.getChildren().add(new Label("Loading..."));

        VBox vBox = new VBox();
        vBox.setSpacing(10);
        vBox.setPadding(new Insets(10, 10, 10, 10));
        vBox.getChildren().addAll(selectFileButton, fileLabel, progressBar, loadingPane);

        primaryStage.setScene(new Scene(vBox, 400, 250));
        primaryStage.show();
    }

    private void processFile(File sourceFile) {
        progressBar.setVisible(true);
        loadingPane.setVisible(true);
        progressBar.setProgress(0);

        Task<Void> task = new Task<Void>() {
            @Override
            protected Void call() throws Exception {
                String sourceFileName = sourceFile.getName();
                String sourceFilePrefixName = StringUtils.substringBeforeLast(sourceFileName, ".");
                String canonicalPath = sourceFile.getCanonicalPath();
                String sourceRootPath = StringUtils.substringBeforeLast(canonicalPath, File.separator);
                String targetRootPath = sourceRootPath + File.separator + sourceFilePrefixName + "_target";
                File sourceRootFile = new File(targetRootPath);
                if (!sourceRootFile.exists()) {
                    sourceRootFile.mkdirs();
                }
                FileUtil.uncompressAllFile(sourceFile, targetRootPath);
                return null;
            }
        };

        task.setOnSucceeded(new EventHandler<WorkerStateEvent>() {
            @Override
            public void handle(WorkerStateEvent event) {
                progressBar.setVisible(false);
                loadingPane.setVisible(false);
                displayResult("文件处理成功!");
                progressBar.progressProperty().unbind();
                progressBar.setProgress(1);
            }
        });
        task.setOnFailed(new EventHandler<WorkerStateEvent>() {
            @Override
            public void handle(WorkerStateEvent event) {
                progressBar.setVisible(false);
                loadingPane.setVisible(false);
                displayResult("文件处理失败!");
                progressBar.progressProperty().unbind();
                progressBar.setProgress(1);
            }
        });

        progressBar.progressProperty().bind(task.progressProperty());

        new Thread(task).start();
    }

    private void displayResult(String result) {
        Stage dialog = new Stage();
        dialog.setTitle("处理结果");

        Label label = new Label(result);
        label.setStyle("-fx-font-size: 14pt; -fx-padding: 20px; -fx-text-alignment: center;");

        StackPane pane = new StackPane();
        pane.getChildren().add(label);

        Scene scene = new Scene(pane, 300, 100);
        dialog.setScene(scene);
        dialog.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
