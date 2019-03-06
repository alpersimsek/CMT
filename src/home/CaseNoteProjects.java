package home;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.input.Clipboard;
import javafx.scene.input.MouseEvent;
import javafx.stage.Stage;

import java.io.*;
import java.net.URL;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.ResourceBundle;
import java.util.Scanner;

public class CaseNoteProjects implements Initializable {

    @FXML
    private TextField txtPrjNoteNum;

    @FXML
    private TextArea txtPrjNote;

    @FXML
    private Button btnPrjNoteClose;

    @FXML
    private TextField txtPrjNoteSeverity;
    @FXML
    private TextField txtPrjNoteAccount;

    @FXML
    private TextField txtPrjNoteSubject;

    @FXML
    private Button btnPrjNoteSave;

    @FXML
    private Button btnPrjNoteClear;

    DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm");


    @FXML
    void handleMouseClicked(MouseEvent event) {

        if (event.getSource() == btnPrjNoteSave){

            if (!txtPrjNote.getText().isEmpty()){

                try {

                    File caseNoteFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\Project\\" + txtPrjNoteNum.getText());

                    if (!caseNoteFile.exists()){
                        new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\Project").mkdir();

                        FileWriter writer = new FileWriter(caseNoteFile);
                        writer.write("====================="+ "\n" + LocalTime.now().format(dtf) + "           " + LocalDate.now() + "\n" + "\n" +
                                txtPrjNote.getText() + "\n" + "\n");

                        writer.close();

                    }else {

                        FileWriter writer = new FileWriter(caseNoteFile, true);
                        writer.append("====================="+"\n" +LocalTime.now().format(dtf) + "           " + LocalDate.now() + "\n" + "\n" + txtPrjNote.getText() + "\n" + "\n");
                        writer.close();
                    }

                }catch (Exception e){
                    e.printStackTrace();
                }

                ((Stage)(btnPrjNoteSave).getScene().getWindow()).close();
            }

        }
        if (event.getSource() == btnPrjNoteClose){

            ((Stage)(btnPrjNoteClose).getScene().getWindow()).close();
        }
        if (event.getSource() == btnPrjNoteClear){
            txtPrjNote.clear();
        }
    }

    private void setFields(){

        File casesel = new File(System.getProperty("user.home") + "\\Documents\\CMT\\CaseDetails\\" + "caseSelProject");

        StringBuilder sb = new StringBuilder();
        BufferedReader br = null;

        try {
            br = new BufferedReader(new FileReader(casesel));
            String line;
            while ((line = br.readLine()) != null) {
                if (sb.length() > 0) {
                    sb.append("\n");
                }
                sb.append(line);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (br != null) {
                    br.close();
                }
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
        String contents = sb.toString();
        //contents = contents.replace("\n", "");
        String[] contentArray = contents.split("\",\"");

        txtPrjNoteNum.setText(contentArray[0]);
        txtPrjNoteSeverity.setText(contentArray[2]);
        txtPrjNoteAccount.setText(contentArray[12]);
        txtPrjNoteSubject.setText(contentArray[15]);

        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                Stage stage = (Stage) txtPrjNoteNum.getScene().getWindow();
                stage.setTitle(txtPrjNoteNum.getText() +  " : PERSONAL NOTE" );
                txtPrjNote.requestFocus();
            }
        });
    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        setFields();
    }
}