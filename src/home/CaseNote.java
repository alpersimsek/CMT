package home;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.MouseEvent;
import javafx.stage.Stage;

import java.io.File;
import java.io.FileWriter;
import java.net.URL;
import java.time.LocalDate;
import java.time.LocalTime;
import java.util.ResourceBundle;

public class CaseNote implements Initializable {

    @FXML
    private TextField txtCaseNoteNum;

    @FXML
    private TextArea txtCaseNote;

    @FXML
    private Button btnCaseNoteClose;

    @FXML
    private TextField txtCaseNoteSeverity;

    @FXML
    private TextField txtCaseNoteSubject;

    @FXML
    private Button btnCaseNoteSave;

    @FXML
    private Button btnCaseNoteClear;

    @FXML
    void handleMouseClicked(MouseEvent event) {

        if (event.getSource() == btnCaseNoteSave){

            if (!txtCaseNote.getText().isEmpty()){

                try {

                    File caseNoteFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\" + txtCaseNoteNum.getText());

                    if (!caseNoteFile.exists()){
                        System.out.println("1");
                        new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes").mkdir();

                        FileWriter writer = new FileWriter(caseNoteFile);
                        writer.write("====================="+ "\n" + LocalTime.now() + "           " + LocalDate.now() + "\n" + "\n" + txtCaseNote.getText() + "\n" + "\n");
                        writer.close();
                        System.out.println("2");

                    }else {

                        System.out.println("3");
                        FileWriter writer = new FileWriter(caseNoteFile, true);
                        writer.append("====================="+"\n" +LocalTime.now() + "           " + LocalDate.now() + "\n" + "\n" + txtCaseNote.getText() + "\n" + "\n");
                        writer.close();
                    }

                }catch (Exception e){
                    e.printStackTrace();
                }

                ((Stage)(btnCaseNoteSave).getScene().getWindow()).close();

            }

        }
        if (event.getSource() == btnCaseNoteClose){

            ((Stage)(btnCaseNoteClose).getScene().getWindow()).close();
        }
        if (event.getSource() == btnCaseNoteClear){
            txtCaseNote.clear();
        }

    }


    @Override
    public void initialize(URL location, ResourceBundle resources) {
        txtCaseNoteNum.setText(Clipboard.getSystemClipboard().getString());
        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                txtCaseNote.requestFocus();
            }
        });
    }
}

