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

import java.io.File;
import java.io.FileWriter;
import java.net.URL;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.ResourceBundle;
import java.util.Scanner;

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
    private TextField txtCaseNoteAccount;

    @FXML
    private TextField txtCaseNoteSubject;

    @FXML
    private Button btnCaseNoteSave;

    @FXML
    private Button btnCaseNoteClear;

    ArrayList<String> caseSelection;

    DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm");


    @FXML
    void handleMouseClicked(MouseEvent event) {

        if (event.getSource() == btnCaseNoteSave){

            if (!txtCaseNote.getText().isEmpty()){

                try {

                    File caseNoteFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\" + txtCaseNoteNum.getText());

                    if (!caseNoteFile.exists()){
                        new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes").mkdir();

                        FileWriter writer = new FileWriter(caseNoteFile);
                        writer.write("====================="+ "\n" + LocalTime.now().format(dtf) + "           " + LocalDate.now() + "\n" + "\n" +
                                txtCaseNote.getText() + "\n" + "\n");

                        writer.close();

                    }else {

                        FileWriter writer = new FileWriter(caseNoteFile, true);
                        writer.append("====================="+"\n" +LocalTime.now().format(dtf) + "           " + LocalDate.now() + "\n" + "\n" + txtCaseNote.getText() + "\n" + "\n");
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

    private void setCaseNumber(){
        txtCaseNoteNum.setText(Clipboard.getSystemClipboard().getString());
    }

    private void setFields(){
        caseSelection = new ArrayList<>();
        File casesel = new File(System.getProperty("user.home") + "\\Documents\\CMT\\" + "caseSel");

        if (casesel.isFile()) {
            Scanner s = null;
            try {
                s = new Scanner(casesel);
            } catch (Exception e) {
                e.printStackTrace();
            }
            while (s.hasNextLine()) {
                caseSelection.add(s.nextLine());
            }

            txtCaseNoteSeverity.setText(caseSelection.get(1));
            txtCaseNoteSubject.setText(caseSelection.get(14));
            txtCaseNoteAccount.setText(caseSelection.get(15));
        }

        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                setHeader();
                txtCaseNote.requestFocus();
            }
        });
    }

    private void setHeader(){

        Stage stage = (Stage) txtCaseNote.getScene().getWindow();
        stage.setTitle(caseSelection.get(0) +  " : PERSONAL NOTE" );

    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        setCaseNumber();
        setFields();
    }
}