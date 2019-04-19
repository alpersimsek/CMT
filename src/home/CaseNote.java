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
    private TextArea txtCaseNotePrev;

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
    String caseNumber;


    @FXML
    void handleMouseClicked(MouseEvent event) {

        if (event.getSource() == btnCaseNoteSave){

            if (!txtCaseNote.getText().isEmpty()){

                try {

                    File caseNoteFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\" + caseNumber);

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

                readNotes(txtCaseNoteNum.getText());
                txtCaseNotePrev.positionCaret(txtCaseNotePrev.getLength());
                txtCaseNote.clear();
                txtCaseNote.requestFocus();

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

        String caseNum = Clipboard.getSystemClipboard().getString();

        caseSelection = new ArrayList<>();
        File casesel = new File(System.getProperty("user.home") + "\\Documents\\CMT\\CaseDetails\\" + caseNum);

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
        }

        txtCaseNoteNum.setText(caseSelection.get(0));
        txtCaseNoteSeverity.setText(caseSelection.get(1));
        txtCaseNoteSubject.setText(caseSelection.get(14));
        txtCaseNoteAccount.setText(caseSelection.get(15));

        readNotes(caseNum);

        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                setHeader();
                txtCaseNote.requestFocus();
            }
        });
    }

    private void setHeader(){

        caseNumber = Clipboard.getSystemClipboard().getString();

        Stage stage = (Stage) txtCaseNote.getScene().getWindow();
        stage.setTitle(caseNumber +  " : MEMO ENTRY" );
    }

    private void readNotes(String str){

        txtCaseNotePrev.clear();

        File prjCase = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\" + str);

        if (prjCase.isFile()) {
            Scanner s = null;
            try {
                s = new Scanner(prjCase);
            } catch (Exception e) {
                e.printStackTrace();
            }
            while (s.hasNextLine()) {
                txtCaseNotePrev.appendText(s.nextLine() + "\n");
            }
            s.close();
        }
    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        //setCaseNumber();
        setFields();
    }
}