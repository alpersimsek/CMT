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

        Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        int screenHeight = screenSize.height;

        FXMLLoader loader;

        if (screenHeight > 1025) {

            loader = new FXMLLoader(getClass().getResource("Main.fxml"));
            Parent root = loader.load();
            primaryStage.setTitle("RBBN Case Management Tool Version 1.12");
            primaryStage.getIcons().add(new Image("home/image/rbbicon.png"));
            primaryStage.setScene(new Scene(root, 1280, 950));
            primaryStage.show();
            primaryStage.setMinHeight(950);
            primaryStage.setMinWidth(1280);
            primaryStage.setMaxWidth(1280);
            primaryStage.setMaxHeight(950);

        } if (screenHeight < 1025) {

            loader = new FXMLLoader(getClass().getResource("Main_2.fxml"));
            Parent root = loader.load();
            primaryStage.setTitle("RBBN Case Management Tool Version 1.12");
            primaryStage.getIcons().add(new Image("home/image/rbbicon.png"));
            primaryStage.setScene(new Scene(root, 1000, 680));
            primaryStage.show();
            primaryStage.setMinHeight(720);
            primaryStage.setMinWidth(1020);
            primaryStage.setMaxHeight(720);
            primaryStage.setMaxWidth(1020);
        }
    }

    public static void main(String[] args) {

       launch(args);
    }
}
