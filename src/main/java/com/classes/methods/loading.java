package com.classes.methods;

import javafx.geometry.Rectangle2D;
import javafx.scene.Group;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.Pane;
import javafx.scene.paint.Color;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

public class loading {
    final Stage initStage = new Stage();

    public void loader() {
        ImageView loadGifView = new ImageView(new Image(getClass().getClassLoader().getResourceAsStream("media/loader.gif")));
        Pane splashLayout = new Pane();
        splashLayout.getChildren().add(loadGifView);
        Group group = new Group();
        group.getChildren().add(splashLayout);
         group.setStyle("-fx-background-color: transparent");
        Scene successScene = new Scene(group, 256, 256);
        successScene.setFill(Color.TRANSPARENT);
        initStage.setTitle("Cargando");
        initStage.getIcons().add(new Image(getClass().getClassLoader().getResourceAsStream("media/icon.png")));
        initStage.setScene(successScene);
        initStage.initStyle(StageStyle.TRANSPARENT);
        initStage.show();

        Rectangle2D primScreenBounds = Screen.getPrimary().getVisualBounds();
        initStage.setX((primScreenBounds.getWidth() - initStage.getWidth()) / 2);
        initStage.setY((primScreenBounds.getHeight() - initStage.getHeight()) / 2);
    }
    public void closeLoader() {
       initStage.close();
    }
}
