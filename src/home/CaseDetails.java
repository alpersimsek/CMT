package home;

import javafx.fxml.FXML;
import javafx.fxml.Initializable;
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
    private Button btnPrjAddNote;

    @FXML
    private Button btnPrjClose;


    File repo = new File(System.getProperty("user.home") + "\\Documents\\CMT\\CaseDetails\\caseSelProject");
    ArrayList<String> prjCaseDetails = new ArrayList<String>();
    ArrayList<String> caseCommentArray;

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
        //contents = contents.replace("\n", "");
        String[] contentArray = contents.split("\",\"");

        txtPrjCaseNum.setText(contentArray[0]);
        txtPrjCaseStat.setText(contentArray[1]);
        txtPrjCaseSev.setText(contentArray[2]);
        txtPrjCaseOwner.setText(contentArray[3]);
        txtPrjAge.setText(contentArray[4]);
        txtPrjProd.setText(contentArray[5]);
        txtPrjHotListR.setText(contentArray[7]);
        txtHotListComm.setText(contentArray[8]);
        txtPrjHotListB.setText(contentArray[9]);
        txtPrjHotListD.setText(contentArray[10]);
        txtPrjGateDate.setText(contentArray[11]);
        txtPrjAcc.setText(contentArray[12]);
        txtPrjReg.setText(contentArray[13]);
        txtPrjCaseSub.setText(contentArray[15]);

        projectCaseComments();
    }


    @FXML
    void handlePrjMouse(MouseEvent event) {

        if (event.getSource() == btnPrjClose){
            ((Stage)(btnPrjClose).getScene().getWindow()).close();
        }
        if (event.getSource() == btnPrjAddNote){
            addNewNote();
        }
    }

    private void addNewNote(){



    }

    private void projectCaseComments(){

        try(HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_comments.xls")))){

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


    @Override
    public void initialize(URL location, ResourceBundle resources) {
        readDetails();

    }
}
