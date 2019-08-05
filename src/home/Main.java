package home;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
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

        if (screenHeight > 1025) {

            loader = new FXMLLoader(getClass().getResource("Main.fxml"));
            Parent root = loader.load();
            primaryStage.setTitle("RBBN Case Management Tool Version 1.13");
            primaryStage.getIcons().add(new Image("home/image/rbbicon.png"));
            primaryStage.setScene(new Scene(root, 1280, 950));
            primaryStage.show();
            primaryStage.setMinHeight(950);
            primaryStage.setMinWidth(1280);
            primaryStage.setMaxWidth(screenWidth);
            primaryStage.setMaxHeight(screenHeight);

        } if (screenHeight < 1025) {

            loader = new FXMLLoader(getClass().getResource("Main_2.fxml"));
            Parent root = loader.load();
            primaryStage.setTitle("RBBN Case Management Tool Version 1.13");
            primaryStage.getIcons().add(new Image("home/image/rbbicon.png"));
            primaryStage.setScene(new Scene(root, 1000, 680));
            primaryStage.show();
            primaryStage.setMinHeight(720);
            primaryStage.setMinWidth(1020);
            primaryStage.setMaxHeight(screenHeight);
            primaryStage.setMaxWidth(screenWidth);
       }
    }

    public static void main(String[] args) {

       launch(args);
    }
}
