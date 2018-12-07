package home;

import javafx.application.Application;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.fxml.FXMLLoader;
import javafx.scene.Group;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.scene.layout.Pane;
import javafx.scene.transform.Scale;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

import java.awt.*;

import static com.sun.org.apache.xerces.internal.utils.SecuritySupport.getResourceAsStream;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception {

        Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        int screenHeight = screenSize.height;

        FXMLLoader loader;

        //loader = new FXMLLoader(getClass().getResource("Main.fxml"));


        if (screenHeight > 1025) {

            loader = new FXMLLoader(getClass().getResource("Main.fxml"));
            Pane root = (Pane) loader.load();
            primaryStage.setTitle("RBBN Case Management Tool");
            primaryStage.getIcons().add(new Image("home/image/rbbicon.png"));
            primaryStage.setScene(new Scene(root, 1280, 950));
            primaryStage.show();


        } else {

            loader = new FXMLLoader(getClass().getResource("Main_2.fxml"));
            Pane root = (Pane) loader.load();
            primaryStage.setTitle("RBBN Case Management Tool");
            primaryStage.getIcons().add(new Image("home/image/rbbicon.png"));
            primaryStage.setScene(new Scene(root, 1000, 680));
            primaryStage.show();
            //primaryStage.setFullScreen(true);
        }
    }

    public static void main(String[] args) {
        launch(args);
    }
}
