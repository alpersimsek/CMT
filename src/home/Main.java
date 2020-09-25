package home;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.image.Image;
import javafx.stage.Stage;
import java.awt.*;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception {

        // GET the screen dimensions of the user
        Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        int screenHeight = screenSize.height;
        int screenWidth = screenSize.width;

        //Load the page depending on the resolution
        FXMLLoader loader;

        if (screenHeight >= 1024) {

            loader = new FXMLLoader(getClass().getResource("Main.fxml"));
            Parent root = loader.load();
            primaryStage.setTitle("RBBN Case Management Tool Version 2.0.1");
            primaryStage.getIcons().add(new Image("home/image/rbbicon.png"));
            primaryStage.setScene(new Scene(root, 1280, 1000));
            primaryStage.show();
            primaryStage.setHeight(980);
            primaryStage.setMinHeight(980);
            primaryStage.setWidth(1110);
            primaryStage.setMinWidth(1110);
            primaryStage.setMaxWidth(screenWidth);
            primaryStage.setMaxHeight(screenHeight);

        } if (screenHeight < 1024) {

            alert("In order to use this program your screen resolution " +
                    "should be at least 1024 pixels in height!" + "\n" +
                    "For instance:  " +
                    "1280 x 1024");

            /*
            loader = new FXMLLoader(getClass().getResource("Main_2.fxml"));
            Parent root = loader.load();
            primaryStage.setTitle("RBBN Case Management Tool Version 1.16.1");
            primaryStage.getIcons().add(new Image("home/image/rbbicon.png"));
            primaryStage.setScene(new Scene(root, 1000, 680));
            primaryStage.show();
            primaryStage.setMinHeight(720);
            primaryStage.setMinWidth(1020);
            primaryStage.setMaxHeight(screenHeight);
            primaryStage.setMaxWidth(screenWidth);
            */
       }
    }

    private void alert(String str){

        Alert alert = new Alert(Alert.AlertType.WARNING);
        ((Stage) alert.getDialogPane().getScene().getWindow()).getIcons().add(new Image("home/image/rbbicon.png"));
        alert.setTitle("RBBN Support Dashboard WARNING:");
        alert.setHeaderText(null);
        alert.setContentText(str);
        alert.showAndWait();
    }

    public static void main(String[] args) {

       launch(args);
    }
}
