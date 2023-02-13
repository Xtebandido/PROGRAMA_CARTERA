package com.app;

import com.classes.connection.conexion;
import com.classes.methods.*;

import javafx.application.Application;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Rectangle2D;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.shape.Rectangle;
import javafx.scene.text.Font;
import javafx.stage.FileChooser;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.event.EventHandler;
import java.io.File;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;

public class app extends Application {

    public static void main(String[] args) {
        launch(args);
    }

    public void start(Stage primaryStage) {
        BorderPane mainPanel = new BorderPane();
        //components
        Button minimize = new Button();
        minimize.setStyle("-fx-background-color: #E8E01E; -fx-background-radius: 1em; -fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #ffc02b; -fx-border-width: 2.5px; -fx-border-radius: 15px");
        Button close = new Button();
        close.setStyle("-fx-background-color: #E81E39; -fx-background-radius: 1em;-fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #ff6158; -fx-border-width: 2.5px; -fx-border-radius: 15px");
        //action
        minimize.setOnAction(event -> {primaryStage.setIconified(true);});
        close.setOnAction(event -> {primaryStage.close(); System.exit(0);});
        //mouse event
        EventHandler<javafx.scene.input.MouseEvent> eventHandler = new EventHandler<javafx.scene.input.MouseEvent>() {
            @Override
            public void handle(MouseEvent e) {
                if (e.getTarget() == minimize) {
                    if (e.getEventType() == MouseEvent.MOUSE_ENTERED) {
                        minimize.setStyle("-fx-background-color: #ffc02b; -fx-background-radius: 1em; -fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #E8E01E; -fx-border-width: 2.5px; -fx-border-radius: 15px");
                    } else {
                        minimize.setStyle("-fx-background-color: #E8E01E; -fx-background-radius: 1em; -fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #ffc02b; -fx-border-width: 2.5px; -fx-border-radius: 15px");
                    }
                } else if (e.getTarget() == close) {
                    if (e.getEventType() == MouseEvent.MOUSE_ENTERED) {
                        close.setStyle("-fx-background-color: #ff6158; -fx-background-radius: 1em; -fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #E81E39; -fx-border-width: 2.5px; -fx-border-radius: 15px");
                    } else {
                        close.setStyle("-fx-background-color: #E81E39; -fx-background-radius: 1em; -fx-min-width: 20px; -fx-max-width: 20px; -fx-min-height: 20px; -fx-max-height: 20px; -fx-border-color: #ff6158; -fx-border-width: 2.5px; -fx-border-radius: 15px");
                    }
                }
            }
        };
        //interaction event
        minimize.addEventFilter(MouseEvent.MOUSE_ENTERED, eventHandler);
        minimize.addEventFilter(MouseEvent.MOUSE_EXITED, eventHandler);
        close.addEventFilter(MouseEvent.MOUSE_ENTERED, eventHandler);
        close.addEventFilter(MouseEvent.MOUSE_EXITED, eventHandler);
        //frame -> top
        ImageView iconView = new ImageView(new Image(getClass().getClassLoader().getResourceAsStream("media/iconTOP.png")));
        HBox buttonsTop = new HBox();
        buttonsTop.getChildren().addAll(minimize, close);
        buttonsTop.setSpacing(10);
        HBox layoutTop = new HBox();
        layoutTop.setPadding(new Insets(15, 15, 15, 15));
        layoutTop.setStyle("-fx-background-color: #07a1e9; -fx-background-radius: 8px;");
        layoutTop.getChildren().add(iconView);
        layoutTop.setSpacing(212);
        layoutTop.getChildren().add(buttonsTop);
        mainPanel.setTop(layoutTop);
        //frame -> center
        Pane layout = new Pane();
        Label selectLabel = new Label("SELECCIONE UN ACTA");
        selectLabel.setFont(new Font("Cooper Black", 26));
        selectLabel.setTextFill(Color.web("#FFFFFF"));
        selectLabel.relocate(75,55);
        layout.getChildren().add(selectLabel);

        TextField tf = new TextField(null);
        tf.setEditable(false);
        tf.setPrefSize(300,10);
        tf.relocate(55,105);
        layout.getChildren().add(tf);

        Button selectButton = new Button("\uD83D\uDDC0");
        selectButton.setTooltip(new Tooltip("Seleccionar"));
        selectButton.relocate(355,105);
        layout.getChildren().add(selectButton);

        final File[] file = {null, null};
        selectButton.setOnAction(event -> {
            FileChooser fileChooser = new FileChooser();
            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Excel", "*.xlsx", "*.xls");
            fileChooser.getExtensionFilters().add(extFilter);
            file[0] = fileChooser.showOpenDialog(null);
            if (file[0] != null) {
                file[1] = file[0];
                tf.setText(file[0].getName());
            }
        });

        Button uploadButton = new Button("▲");
        uploadButton.setTooltip(new Tooltip("Subir"));
        uploadButton.relocate(385,105);
        layout.getChildren().add(uploadButton);
        uploadButton.setOnAction(event -> {
            if (tf.getText() != null && tf.getText() != "") {
                primaryStage.close();

                Stage initStage = new Stage();
                new loading(initStage);

                uploadXLS classUpload = new uploadXLS();
                try {
                    new Thread(() -> {
                        classUpload.upload(file[1], initStage, tf);
                    }).start();
                } catch (Exception e) {
                    Alert alert = new Alert(Alert.AlertType.WARNING);
                    alert.setHeaderText(null);
                    alert.setTitle("Error");
                    alert.setContentText("ERROR DESCONOCIDO.\n\nLog: " + e);
                    alert.showAndWait();
                }

            } else {
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.setHeaderText(null);
                alert.setTitle("Error");
                alert.setContentText("SELECCIONE UN ARCHIVO.");
                alert.showAndWait();
            }
        });

        Label filterLabel = new Label("FILTRE UNA FECHA");
        filterLabel.setFont(new Font("Cooper Black", 26));
        filterLabel.setTextFill(Color.web("#FFFFFF"));
        filterLabel.relocate(100,150);
        layout.getChildren().add(filterLabel);

        ObservableList<String> optionsMonth = FXCollections.observableArrayList("1 - ENERO", "2 - FEBRERO", "3 - MARZO", "4 - ABRIL", "5 - MAYO", "6 - JUNIO", "7 - JULIO", "8 - AGOSTO", "9 - SEPTIEMBRE", "10 - OCTUBRE", "11 - NOVIEMBRE", "12 - DICIEMBRE");
        final ComboBox comboMonth = new ComboBox(optionsMonth);
        comboMonth.setPrefSize(175,30);
        comboMonth.relocate(55,195);
        comboMonth.setPromptText("SELECCIONAR MES");
        layout.getChildren().add(comboMonth);

        conexion database = new conexion();
        Connection con = database.conectarSQL();

        List<String> years = new ArrayList<>();
        try {
            PreparedStatement ps = con.prepareStatement("SELECT DISTINCT strftime('%Y', f_cierre) f_cierre FROM IMPRESION ORDER BY f_cierre DESC");
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                String year = rs.getString("f_cierre");
                years.add(year);
            }
        } catch (Exception e) {
            System.out.println(e);
        }

        ObservableList<String> optionsYear = FXCollections.observableArrayList(new ArrayList<>(years));
        final ComboBox comboYear = new ComboBox(optionsYear);
        comboYear.setPrefSize(177,30);
        comboYear.relocate(235,195);
        comboYear.setPromptText("SELECCIONAR AÑO");
        layout.getChildren().add(comboYear);

        Button generateButton = new Button("GENERAR INFORME");
        generateButton.relocate(55,260);
        generateButton.setPrefSize(360,50);
        layout.getChildren().add(generateButton);
        generateButton.setOnAction(event -> {
            if (comboMonth.getSelectionModel().getSelectedIndex() != -1 && comboYear.getSelectionModel().getSelectedIndex() != -1) {
                primaryStage.close();

                Stage initStage = new Stage();
                new loading(initStage);

                new Thread (() -> {
                    generateXLS gen = new generateXLS();
                    gen.generate(initStage, comboMonth.getValue().toString(), comboYear.getValue().toString());
                }).start();
            } else {
                if (comboMonth.getSelectionModel().getSelectedIndex() == -1 && comboYear.getSelectionModel().getSelectedIndex() != -1) {
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setHeaderText(null);
                    alert.setTitle("Error");
                    alert.setContentText("NO SE HA SELECCIONADO UN MES.");
                    alert.showAndWait();
                } else if (comboMonth.getSelectionModel().getSelectedIndex() != -1 && comboYear.getSelectionModel().getSelectedIndex() == -1) {
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setHeaderText(null);
                    alert.setTitle("Error");
                    alert.setContentText("NO SE HA SELECCIONADO UN AÑO.");
                    alert.showAndWait();
                } else {
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setHeaderText(null);
                    alert.setTitle("Error");
                    alert.setContentText("NO SE HA SELECCIONADO MES Y AÑO.");
                    alert.showAndWait();
                }
            }
        });

        Button historicButton = new Button("GENERAR HISTORICO");
        historicButton.relocate(55,320);
        historicButton.setPrefSize(360,50);
        layout.getChildren().add(historicButton);
        historicButton.setOnAction(event -> {
            primaryStage.close();

            Stage initStage = new Stage();
            new loading(initStage);

            new Thread (() -> {
                historic h = new historic();
                h.historic(initStage);
            }).start();
        });

        mainPanel.setCenter(layout);
        //root
        GridPane root = new GridPane();
        root.setStyle("-fx-background-color: #4EBAEC; -fx-border-color: #000000; -fx-border-width: 5px; -fx-border-radius: 8px");
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