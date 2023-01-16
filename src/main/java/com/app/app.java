package com.app;

import com.classes.methods.leerExcel;
import com.classes.methods.loading;
import javafx.application.Application;

import javafx.geometry.Insets;
import javafx.geometry.Rectangle2D;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.shape.Rectangle;
import javafx.scene.text.Font;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.event.EventHandler;
import java.io.File;

public class app extends Application {

    public static void main(String[] args) {
        launch(args);
    }

    public void start(Stage primaryStage) {
        BorderPane mainPanel = new BorderPane();
        //components
        Button minimize = new Button();
        minimize.setStyle("-fx-background-color: #ffc02b; -fx-background-radius: 1em; -fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #ffc02b; -fx-border-width: 2.5px; -fx-border-radius: 15px");
        Button close = new Button();
        close.setStyle("-fx-background-color: #ff6158; -fx-background-radius: 1em;-fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #ff6158; -fx-border-width: 2.5px; -fx-border-radius: 15px");
        //action
        minimize.setOnAction(event -> {primaryStage.setIconified(true);});
        close.setOnAction(event -> {primaryStage.close(); System.exit(0);});
        //mouse event
        EventHandler<javafx.scene.input.MouseEvent> eventHandler = new EventHandler<javafx.scene.input.MouseEvent>() {
            @Override
            public void handle(MouseEvent e) {
                if (e.getTarget() == minimize) {
                    if (e.getEventType() == MouseEvent.MOUSE_ENTERED) {
                        minimize.setStyle("-fx-background-color: #FFFFFF; -fx-background-radius: 1em; -fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #ffc02b; -fx-border-width: 2.5px; -fx-border-radius: 15px");
                    } else {
                        minimize.setStyle("-fx-background-color: #ffc02b; -fx-background-radius: 1em; -fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #ffc02b; -fx-border-width: 2.5px; -fx-border-radius: 15px");
                    }
                } else if (e.getTarget() == close) {
                    if (e.getEventType() == MouseEvent.MOUSE_ENTERED) {
                        close.setStyle("-fx-background-color: #FFFFFF; -fx-background-radius: 1em; -fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #ff6158; -fx-border-width: 2.5px; -fx-border-radius: 15px");
                    } else {
                        close.setStyle("-fx-background-color: #ff6158; -fx-background-radius: 1em; -fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #ff6158; -fx-border-width: 2.5px; -fx-border-radius: 15px");
                    }
                }
            }
        };
        //interaction event
        minimize.addEventFilter(MouseEvent.MOUSE_ENTERED, eventHandler);
        minimize.addEventFilter(MouseEvent.MOUSE_EXITED, eventHandler);
        close.addEventFilter(MouseEvent.MOUSE_ENTERED, eventHandler);
        close.addEventFilter(MouseEvent.MOUSE_EXITED, eventHandler);
        //frame
        //top
        ImageView iconView = new ImageView(new Image(getClass().getClassLoader().getResourceAsStream("media/iconTOP.png")));
        HBox buttonsTop = new HBox();
        buttonsTop.getChildren().addAll(minimize, close);
        buttonsTop.setSpacing(10);
        HBox layoutTop = new HBox();
        layoutTop.setPadding(new Insets(15, 30, 15, 12));
        layoutTop.setStyle("-fx-background-color: #07a1e9;");
        layoutTop.getChildren().add(iconView);
        layoutTop.setSpacing(230);
        layoutTop.getChildren().add(buttonsTop);
        mainPanel.setTop(layoutTop);
        //center
        Pane layout = new Pane();
        Text selectLabel = new Text("SELECCIONE UN ACTA");
        selectLabel.setFont(new Font("OCR A", 26));
        selectLabel.relocate(105,55);
        layout.getChildren().add(selectLabel);

        TextField tf = new TextField();
        tf.setEditable(false);
        tf.setPrefSize(300,10);
        tf.relocate(55,105);
        layout.getChildren().add(tf);

        Button selectButton = new Button("\uD83D\uDDC0");
        selectButton.setTooltip(new Tooltip("Seleccionar"));
        selectButton.relocate(355,105);
        layout.getChildren().add(selectButton);

        final File[] file = {null};

        selectButton.setOnAction(event -> {
            FileChooser fileChooser = new FileChooser();
            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Excel", "*.xlsx", "*.xls");
            fileChooser.getExtensionFilters().add(extFilter);

            file[0] = fileChooser.showOpenDialog(null);
            if (file[0] != null) {
                tf.setText(file[0].getName());
            }
        });

        Button uploadButton = new Button("▲");
        uploadButton.setTooltip(new Tooltip("Subir"));
        uploadButton.relocate(385,105);
        layout.getChildren().add(uploadButton);
        uploadButton.setOnAction(event -> {
            if (file[0] != null) {
                leerExcel xlsx = new leerExcel();
                xlsx.leerExcel(file[0], primaryStage);
            } else {
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.setTitle(null);
                alert.setHeaderText(null);
                alert.setContentText("ERROR: SELECCIONE UN ARCHIVO");
                alert.showAndWait();
            }
        });

        mainPanel.setCenter(layout);
        //root
        GridPane root = new GridPane();
        root.getChildren().add(mainPanel);
        //border root
        Rectangle rect = new Rectangle(0,0,500,500);
        rect.setArcHeight(30);
        rect.setArcWidth(30);
        root.setClip(rect);
        //primaryStage
        primaryStage.setScene(new Scene(root, 500, 500, Color.TRANSPARENT));
        primaryStage.setTitle("Acueducto");
        primaryStage.getIcons().add(new Image(getClass().getClassLoader().getResourceAsStream("media/icon.png")));
        primaryStage.initStyle(StageStyle.TRANSPARENT);
        primaryStage.show();
        //center stage
        Rectangle2D primScreenBounds = Screen.getPrimary().getVisualBounds();
        primaryStage.setX((primScreenBounds.getWidth() - primaryStage.getWidth()) / 2);
        primaryStage.setY((primScreenBounds.getHeight() - primaryStage.getHeight()) / 2);
    }

}