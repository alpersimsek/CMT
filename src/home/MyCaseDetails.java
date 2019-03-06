package home;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.input.Clipboard;
import javafx.scene.input.MouseEvent;
import javafx.stage.Stage;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.ResourceBundle;
import java.util.Scanner;

public class MyCaseDetails implements Initializable {

    @FXML
    private TextField txtMyCaseDetNum;

    @FXML
    private TextField txtMyCaseDetSev;

    @FXML
    private TextField txtMyCaseDetStat;

    @FXML
    private TextField txtMyCaseDetAge;

    @FXML
    private TextField txtMyCaseDetOwn;

    @FXML
    private TextField txtMyCaseDetCoOwn;

    @FXML
    private TextField txtMyCaseDetCoOwnQueue;

    @FXML
    private TextField txtMyCaseDetResp;

    @FXML
    private TextField txtMyCaseDetUpd;

    @FXML
    private TextField txtMyCaseDetProd;

    @FXML
    private TextField txtMyCaseDetAcc;

    @FXML
    private TextField txtMyCaseDetReg;

    @FXML
    private TextField txtMyCaseDetEsc;

    @FXML
    private TextField txtMyCaseDetLevel;

    @FXML
    private TextField txtMyCaseDetType;

    @FXML
    private TextField txtMyCaseDetSec;

    @FXML
    private CheckBox chckMyDetFol;

    @FXML
    private TextField txtMyCaseDetSub;

    @FXML
    private TextArea txtMyCaseDetComm;

    @FXML
    private TextArea txtMyCaseDetNotes;

    @FXML
    private TextArea txtMyCaseDetAddNote;

    @FXML
    private Button btnSaveNote;

    @FXML
    private Button btnClose;

    @FXML
    private Button btnDeletNote;

    ArrayList<String> caseSelection;
    ArrayList<String> caseCommentArray;



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

            txtMyCaseDetNum.setText(caseSelection.get(0));
            txtMyCaseDetSev.setText(caseSelection.get(1));
            txtMyCaseDetStat.setText(caseSelection.get(2));
            txtMyCaseDetOwn.setText(caseSelection.get(3));
            txtMyCaseDetCoOwn.setText(caseSelection.get(4));
            txtMyCaseDetCoOwnQueue.setText(caseSelection.get(5));
            txtMyCaseDetResp.setText(caseSelection.get(6));
            txtMyCaseDetAge.setText(caseSelection.get(7));
            txtMyCaseDetUpd.setText(caseSelection.get(8));
            txtMyCaseDetEsc.setText(caseSelection.get(9));
            txtMyCaseDetLevel.setText(caseSelection.get(10));
            if (caseSelection.get(11).equals("1")){
                chckMyDetFol.setSelected(true);
            }else{
                chckMyDetFol.setSelected(false);
            }
            txtMyCaseDetType.setText(caseSelection.get(12));
            txtMyCaseDetProd.setText(caseSelection.get(13));
            txtMyCaseDetSub.setText(caseSelection.get(14));
            txtMyCaseDetAcc.setText(caseSelection.get(15));
            txtMyCaseDetReg.setText(caseSelection.get(16));
            txtMyCaseDetSec.setText(caseSelection.get(17));
            processComments(caseSelection.get(0));
            readNotes(caseSelection.get(0));

        }

        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                setHeader();
                txtMyCaseDetAddNote.requestFocus();
            }
        });

    }

    private void setHeader(){

        Stage stage = (Stage) txtMyCaseDetNum.getScene().getWindow();
        stage.setTitle(caseSelection.get(0) +  " : CASE DETAIL WINDOW" );
    }

    private void readNotes(String str){
        txtMyCaseDetNotes.clear();

        File prjCase = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes\\" + str);

        if (prjCase.isFile()) {
            Scanner s = null;
            try {
                s = new Scanner(prjCase);
            } catch (Exception e) {
                e.printStackTrace();
            }
            while (s.hasNextLine()) {
                txtMyCaseDetNotes.appendText(s.nextLine() + "\n");
            }
            s.close();
        }
    }

    private void processComments(String str1){
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

            String caseNumber = str1;
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

                txtMyCaseDetComm.setText("THERE ARE NO WORK NOTES LOGGED FOR THIS CASE SINCE LAST 7 DAYS!");
            }

            for (int i = 0; i < arraySize; i += 2) {

                txtMyCaseDetComm.appendText("===============" + "\n" + caseCommentArray.get(i)+ "\n" + "\n" + caseCommentArray.get(i+1) + "\n");
            }
            txtMyCaseDetComm.positionCaret(0);

        }catch (Exception e){
            e.printStackTrace();
        }

        caseCommentArray.clear();
    }

    private void addNote(){

    }


    @FXML
    void handleMouseClick(MouseEvent event) {

        if(event.getSource() == btnClose){
            ((Stage)(btnClose).getScene().getWindow()).close();
        }
        if (event.getSource() == btnSaveNote){
            addNote();
        }

    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        setFields();
    }
}
