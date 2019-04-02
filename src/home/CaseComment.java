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

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.ResourceBundle;
import java.util.Scanner;


public class CaseComment implements Initializable {

    @FXML
    private TextArea txtCaseComments;

    @FXML
    private Button btnClose;

    @FXML
    private Button btnNewNote;

    @FXML
    private TextField txtComNum;

    @FXML
    private TextField txtComSev;

    @FXML
    private TextField txtComAcc;

    @FXML
    private TextField txtComSubj;

    ArrayList<String> caseCommentArray;
    ArrayList<String> caseSelection;



    private void setCaseNumber(){

    }

    private void viewCases(){

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
                Alert alert = new Alert(Alert.AlertType.WARNING);
                ((Stage) alert.getDialogPane().getScene().getWindow()).getIcons().add(new Image("home/image/rbbicon.png"));
                alert.setTitle("RBBN CASE MANAGEMENT TOOL WARNING:");
                alert.setHeaderText(null);
                alert.setContentText("THERE IS NO WORK COMMENT FOR THIS CASE"+ "\n" + "SINCE 7 DAYS!");
                alert.showAndWait();
            }

                for (int i = 0; i < arraySize; i += 2) {

                    txtCaseComments.appendText("===============" + "\n" + caseCommentArray.get(i)+ "\n" + "\n" + caseCommentArray.get(i+1) + "\n");
                }
                txtCaseComments.positionCaret(0);

        }catch (Exception e){
            e.printStackTrace();
        }

        caseCommentArray.clear();
    }

    @FXML
    void handleMouseClicked(MouseEvent event) {

        if (event.getSource() == btnClose){

            ((Stage)(btnClose).getScene().getWindow()).close();
        }
        if (event.getSource() == btnNewNote){
            newNote();
        }
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

            txtComNum.setText(caseSelection.get(0));
            txtComSev.setText(caseSelection.get(1));
            txtComSubj.setText(caseSelection.get(14));
            txtComAcc.setText(caseSelection.get(15));
        }

        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                setHeader();
                txtCaseComments.requestFocus();
            }
        });
    }

    private void setHeader(){

        Stage stage = (Stage) txtCaseComments.getScene().getWindow();
        stage.setTitle(caseSelection.get(0) +  " : WORK NOTE(S) FROM LAST 7 DAYS" );

    }

    private void newNote(){

        ((Stage)(btnClose).getScene().getWindow()).close();

        Parent root;
        try {
            root = FXMLLoader.load(getClass().getClassLoader().getResource("home/CaseNote.fxml"));
            Stage stage = new Stage();
            stage.setTitle("ADD MEMO");
            stage.getIcons().add(new Image("home/image/rbbicon.png"));
            stage.setScene(new Scene(root, 650, 400));
            stage.show();
            stage.setMinWidth(650);
            stage.setMinHeight(420);
            stage.setMaxWidth(650);
            stage.setMaxHeight(420);

        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        viewCases();
        setFields();
    }
}
