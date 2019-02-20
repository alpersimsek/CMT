package home;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.input.MouseEvent;
import javafx.scene.text.Text;
import javafx.stage.Stage;

import java.awt.*;
import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

public class Login implements Initializable {

    String password = "rbbnforecast!";

    @FXML
    private TextField txtpass;

    @FXML
    private Button btnLogin;

    @FXML
    private Button btnClose;

    @FXML
    private Text txtWrongPass;

    @FXML
    void handleMouseClicked(MouseEvent event) {

        if(event.getSource() == btnClose){
            ((Stage)(btnClose).getScene().getWindow()).close();
        }

        if (event.getSource() == btnLogin){
            txtWrongPass.setVisible(false);
            checkUser();
        }
        if (event.getSource() == txtpass){
            txtWrongPass.setVisible(false);
            txtpass.clear();
        }
    }

    private void checkUser(){

        if (!txtpass.getText().equals("")){

            String promptedpass = txtpass.getText();
            if (promptedpass.equals(password)){
                try {
                    openForecastPage();
                } catch (Exception e) {
                    e.printStackTrace();
                }

                ((Stage)(btnLogin).getScene().getWindow()).close();

            }
            else{
                txtWrongPass.setVisible(true);
            }
        }

    }

    private void openForecastPage () throws Exception{

        FXMLLoader loader;

        Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        int screenHeight = screenSize.height;

        Parent root;
        try {
            root = FXMLLoader.load(getClass().getClassLoader().getResource("home/Model.fxml"));
            Stage stage = new Stage();
            stage.setTitle("RBBN Case Management Tool Forecast Modeling Page");
            stage.getIcons().add(new Image("home/image/rbbicon.png"));
            stage.setScene(new Scene(root, 1024, 768));
            stage.show();
            stage.setMinWidth(1024);
            stage.setMinHeight(768);
            stage.setMaxWidth(1024);
            stage.setMaxHeight(768);

        }
        catch (IOException e) {
            e.printStackTrace();
        }

    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                txtpass.requestFocus();
            }
        });
    }
}
