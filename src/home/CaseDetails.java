package home;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.input.Clipboard;
import javafx.scene.input.MouseEvent;
import javafx.stage.Stage;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.net.URL;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.ResourceBundle;
import java.util.Scanner;

public class CaseDetails implements Initializable {

    @FXML
    private TextField txtPrjCaseNum;
    @FXML
    private TextField txtPrjCaseSev;
    @FXML
    private TextField txtPrjCaseStat;
    @FXML
    private TextField txtPrjCaseOwner;
    @FXML
    private TextField txtPrjHotListR;
    @FXML
    private TextField txtPrjCaseSub;
    @FXML
    private TextField txtPrjHotListB;
    @FXML
    private TextField txtPrjProd;
    @FXML
    private TextField txtPrjHotListD;
    @FXML
    private TextField txtPrjGateDate;
    @FXML
    private TextField txtPrjAcc;
    @FXML
    private TextField txtPrjAge;
    @FXML
    private TextField txtPrjReg;
    @FXML
    private TextArea txtHotListComm;
    @FXML
    private TextArea txtPrjCaseComment;
    @FXML
    private TextArea txtPrjNote;
    @FXML
    private TextArea txtPrjAddNote;
    @FXML
    private Button btnPrjAddNote;
    @FXML
    private Button btnPrjClose;
    @FXML
    private Button btnPrjDelNote;


    File repo = new File(System.getProperty("user.home") + "\\Documents\\CMT\\CaseDetails\\caseSelProject");
    ArrayList<String> prjCaseDetails = new ArrayList<String>();
    ArrayList<String> caseCommentArray;
    String[] contentArray;

    DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm");


    private void readDetails() {

        StringBuilder sb = new StringBuilder();
        BufferedReader br = null;
        try {
            br = new BufferedReader(new FileReader(repo));
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
        contents = contents.replace(".0000000000", "");

        contentArray = contents.split("\",\"");


        txtPrjCaseNum.setText(contentArray[0]);
        txtPrjCaseSev.setText(contentArray[1]);
        txtPrjCaseStat.setText(contentArray[2]);
        txtPrjCaseOwner.setText(contentArray[3]);
        txtPrjHotListR.setText(contentArray[4]);
        txtPrjCaseSub.setText(contentArray[5]);
        txtPrjHotListB.setText(contentArray[6]);
        txtPrjProd.setText(contentArray[7]);
        txtPrjHotListD.setText(contentArray[8]);
        txtPrjGateDate.setText(contentArray[9]);
        txtPrjAcc.setText(contentArray[10]);
        txtPrjAge.setText(contentArray[11]);
        txtPrjReg.setText(contentArray[12]);
        txtHotListComm.setText(contentArray[13]);

        projectCaseComments();
        readNotes(txtPrjCaseNum.getText());

        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                setHeader();
                txtPrjAddNote.requestFocus();
            }
        });
    }


    private void setHeader(){

        Stage stage = (Stage) txtPrjCaseNum.getScene().getWindow();
        stage.setTitle(txtPrjCaseNum.getText() +  " : Case Details..." );
    }

    @FXML
    void handlePrjMouse(MouseEvent event) {

        if (event.getSource() == btnPrjClose){
            ((Stage)(btnPrjClose).getScene().getWindow()).close();
        }
        if (event.getSource() == btnPrjAddNote){
            addNewNote();
        }
        if (event.getSource() == btnPrjDelNote){
            delNote();
        }
    }

    private void addNewNote(){
        if (!txtPrjAddNote.getText().isEmpty()) {

            try {

                File caseNoteFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\Project\\" + txtPrjCaseNum.getText());

                if (!caseNoteFile.exists()) {
                    new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes").mkdir();
                    new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\Project").mkdir();

                    FileWriter writer = new FileWriter(caseNoteFile);
                    writer.write("=====================" + "\n" + LocalTime.now().format(dtf) + "           " + LocalDate.now() + "\n" + "\n" +
                            txtPrjAddNote.getText() + "\n" + "\n");

                    writer.close();

                } else {

                    FileWriter writer = new FileWriter(caseNoteFile, true);
                    writer.append("=====================" + "\n" + LocalTime.now().format(dtf) + "           " + LocalDate.now() + "\n" + "\n" + txtPrjAddNote.getText() + "\n" + "\n");
                    writer.close();
                }

            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        readNotes(txtPrjCaseNum.getText());
        txtPrjAddNote.clear();
    }

    private void delNote(){

        File caseNoteFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\Project\\" + txtPrjCaseNum.getText());
        caseNoteFile.delete();
        readNotes(txtPrjCaseNum.getText());

    }

    private void projectCaseComments(){

        try(HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\Data\\cmt_comments.xls")))){

            caseCommentArray = new ArrayList<>();
            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int mycaseNumCellRef = 0;
            int myCaseCommentDateRef = 0;
            int myCaseCommentRef = 0;
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();

            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            String caseNumber = Clipboard.getSystemClipboard().getString();
            //txtComNum.setText(caseNumber + "  " + "Work Notes From Last 7 Days: ");

            for (int i = 0; i < cellnum ; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Case Number")) {
                    mycaseNumCellRef = i;
                }
                if (filterColName.equals("Work Note: Created Date")) {
                    myCaseCommentDateRef = i;
                }
                if (filterColName.equals("Work Comments")) {
                    myCaseCommentRef = i;
                }
            }

            for (int i = 1; i < lastRow + 1; i++) {

                cellVal1 = filtersheet.getRow(i).getCell(mycaseNumCellRef);
                String commentCaseNumber = cellVal1.getStringCellValue();

                cellVal2 = filtersheet.getRow(i).getCell(myCaseCommentDateRef);
                String commentDate = cellVal2.getStringCellValue();

                cellVal3 = filtersheet.getRow(i).getCell(myCaseCommentRef);
                String commentComment = cellVal3.getStringCellValue();

                if (commentCaseNumber.equals(caseNumber)){

                    caseCommentArray.add(commentDate);
                    caseCommentArray.add(commentComment);
                }
            }

            int arraySize = caseCommentArray.size();

            if (arraySize == 0){

                txtPrjCaseComment.setText("THERE ARE NO WORK NOTES LOGGED FOR THIS CASE SINCE LAST 7 DAYS!");
            }

            for (int i = 0; i < arraySize; i += 2) {

                txtPrjCaseComment.appendText("===============" + "\n" + caseCommentArray.get(i)+ "\n" + "\n" + caseCommentArray.get(i+1) + "\n");
            }
            txtPrjCaseComment.positionCaret(0);

        }catch (Exception e){
            e.printStackTrace();
        }

        caseCommentArray.clear();
    }

    private void readNotes(String str){

        txtPrjNote.clear();
        File prjCase = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\Project\\" + str);

        if (prjCase.isFile()) {
            Scanner s = null;
            try {
                s = new Scanner(prjCase);
            } catch (Exception e) {
                e.printStackTrace();
            }
            while (s.hasNextLine()) {
                txtPrjNote.appendText(s.nextLine() + "\n");
            }
            s.close();
        }
    }


    @Override
    public void initialize(URL location, ResourceBundle resources) {
        readDetails();
    }
}
