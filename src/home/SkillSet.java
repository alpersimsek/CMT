package home;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.event.EventHandler;
import javafx.fxml.Initializable;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.net.URL;
import java.util.*;


public class SkillSet implements Initializable {

    @FXML
    private RadioButton rdEngMyTeam;

    @FXML
    private RadioButton rdEngOverall;

    @FXML
    private ListView<String> engNameList;
    @FXML
    private ListView<String> engNameListAll;

    @FXML
    private ListView<String> engSkilLev;

    @FXML
    private ListView<String> engSkillName;

    @FXML
    private RadioButton rdSkilMyTeam;

    @FXML
    private RadioButton rdSkillOverAll;

    @FXML
    private ListView<String> skillNameList;

    @FXML
    private ListView<String> skillNameListAll;

    @FXML
    private ListView<String> skillLevelList;

    @FXML
    private ListView<String> skillEngName;

    @FXML
    private TextField txtSearchEng;

    @FXML
    private TextField txtSearchSkill;

    @FXML
    private Text lblSearcEng;

    @FXML
    private Text lblSearchSkill;

    @FXML
    private Button btnExport;

    ArrayList<String> readUserList;
    ArrayList<String> readOverAllUsers;
    ArrayList<String> safeUserList;
    int userRef;
    int skillRef;
    String selectedLevel;
    String selected;
    String selectedSkill;
    String selectedSkillLevel;
    ArrayList<String> expertLevel;
    ArrayList<String> intLevel;
    ArrayList<String> basicLevel;
    ArrayList<String> noLevel;
    ArrayList<String> skillsAll;
    ArrayList<String> skillsExpert;
    ArrayList<String> skillsInterm;
    ArrayList<String> skillsBegin;
    ArrayList<String> skillsNone;

    ObservableList<String> levels = FXCollections.observableArrayList();

    @FXML
    void handleRadioClick(MouseEvent event) {

        if(event.getSource() == rdEngMyTeam){
            rdEngMyTeam.setSelected(true);
            rdEngOverall.setSelected(false);
            engNameListAll.setVisible(false);
            engNameList.setVisible(true);
            engSkilLev.getItems().clear();
            engSkillName.getItems().clear();
            engSkilLev.setVisible(false);
            engSkillName.setVisible(false);
            engMyTeam();
        }
        if(event.getSource() == rdEngOverall){
            rdEngMyTeam.setSelected(false);
            rdEngOverall.setSelected(true);
            engNameList.setVisible(false);
            engNameListAll.setVisible(true);
            engSkilLev.getItems().clear();
            engSkillName.getItems().clear();
            engSkilLev.setVisible(false);
            engSkillName.setVisible(false);
            engOverAllTeam();
        }
        if(event.getSource() == rdSkilMyTeam){
            rdSkilMyTeam.setSelected(true);
            rdSkillOverAll.setSelected(false);
            skillNameList.setVisible(true);
            skillNameListAll.setVisible(false);
            skillLevelList.getItems().clear();
            skillEngName.getItems().clear();
            skillLevelList.setVisible(false);
            skillEngName.setVisible(false);
            btnExport.setVisible(false);
            skillMyTeam();
        }
        if (event.getSource() == rdSkillOverAll){
            rdSkilMyTeam.setSelected(false);
            rdSkillOverAll.setSelected(true);
            skillNameList.setVisible(false);
            skillNameListAll.setVisible(true);
            skillNameListAll.getItems().clear();
            skillEngName.getItems().clear();
            skillLevelList.setVisible(false);
            skillEngName.setVisible(false);
            btnExport.setVisible(false);
            skillOverAllTeam();
        }
    }

    private void engMyTeam(){

        txtSearchEng.clear();
        int userarraysize = safeUserList.size();

        ObservableList<String> users = FXCollections.observableArrayList();

        for (int i = 0; i <userarraysize ; i++) {
            users.add(safeUserList.get(i));
        }

        if (engNameList.getItems().size() == 0){
            engNameList.getItems().addAll(users);
        }
        engNameList.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);

        lblSearcEng.setVisible(true);
        txtSearchEng.setVisible(true);

        FilteredList<String> filteredEng = new FilteredList((ObservableList) users, p -> true);

        txtSearchEng.textProperty().addListener((observable, oldValue, newValue) -> {
            filteredEng.setPredicate(string -> {

                engSkilLev.setVisible(false);
                engSkillName.setVisible(false);

                if (newValue == null || newValue.isEmpty()) {
                    return true;
                }

                String lowerCaseCustomerName = newValue.toLowerCase();

                if (string.toLowerCase().contains(lowerCaseCustomerName)) {
                    return true;
                }
                return false;
            });
        });

        engNameList.setItems(filteredEng);

        engNameList.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                engSkillName.getItems().clear();
                selected = "";
                selectedLevel = "";

                selected = engNameList.getSelectionModel().getSelectedItem();
                userRef =0;
                engSkilLev.getItems().clear();
                levels.clear();
                setLevels();
                engSkilLev.setVisible(true);
                engSkillName.setVisible(false);
                engSkilLev.getItems().addAll(levels);

                engSkilLev.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
                engSkilLev.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        engSkillName.getItems().clear();
                        selectedLevel = engSkilLev.getSelectionModel().getSelectedItem();
                        engSkillName.setVisible(true);


                        try {

                            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\SkillSet\\SkillsetProfiles.xls")));
                            HSSFSheet sheet = workbook.getSheetAt(0);
                            HSSFCell cellVal;

                            int cellnum = sheet.getRow(0).getLastCellNum();
                            int lastRow = sheet.getLastRowNum();
                            expertLevel = new ArrayList<>();
                            intLevel = new ArrayList<>();
                            basicLevel = new ArrayList<>();
                            noLevel = new ArrayList<>();

                            for (int i = 1; i <cellnum ; i++) {
                                String userNameColl = sheet.getRow(0).getCell(i).getStringCellValue();
                                if (userNameColl.equals(selected)){
                                    userRef = i;
                                }
                            }

                            for (int i = 1; i <lastRow ; i++) {

                                cellVal = sheet.getRow(i).getCell(userRef);
                                int cellValue = 0;

                                if (cellVal != null) {

                                    if (cellVal.getCellType() == CellType.NUMERIC){
                                        cellValue = (int) cellVal.getNumericCellValue();
                                    }
                                    else{
                                        cellValue = Integer.parseInt(cellVal.getStringCellValue());
                                    }

                                    if (cellValue == 3) {
                                        expertLevel.add(sheet.getRow(i).getCell(0).getStringCellValue());
                                    }
                                    if (cellValue == 2) {
                                        intLevel.add(sheet.getRow(i).getCell(0).getStringCellValue());
                                    }
                                    if (cellValue == 1) {
                                        basicLevel.add(sheet.getRow(i).getCell(0).getStringCellValue());
                                    }
                                    if (cellValue == 0) {
                                        noLevel.add(sheet.getRow(i).getCell(0).getStringCellValue());
                                    }
                                }else {
                                    noLevel.add(sheet.getRow(i).getCell(0).getStringCellValue());
                                }
                            }

                        }catch (Exception e){
                            e.printStackTrace();
                        }

                        int expertSize = expertLevel.size();
                        int intermSize = intLevel.size();
                        int basicSize = basicLevel.size();
                        int noSize = noLevel.size();

                        ObservableList<String> skills = FXCollections.observableArrayList();

                        if (selectedLevel.equals("EXPERT")){

                            for (int i = 0; i <expertSize ; i++) {
                                skills.addAll(expertLevel.get(i));
                            }

                            engSkillName.getItems().addAll(skills);
                        }
                        if (selectedLevel.equals("INTERMEDIATE")){

                            for (int i = 0; i <intermSize ; i++) {
                                skills.addAll(intLevel.get(i));
                            }

                            engSkillName.getItems().addAll(skills);
                        }
                        if (selectedLevel.equals("BEGINNER")){

                            for (int i = 0; i <basicSize ; i++) {
                                skills.addAll(basicLevel.get(i));
                            }

                            engSkillName.getItems().addAll(skills);
                        }
                        if (selectedLevel.equals("NONE")){

                            for (int i = 0; i <noSize ; i++) {
                                skills.addAll(noLevel.get(i));
                            }
                            engSkillName.getItems().addAll(skills);
                        }
                    }
                });
            }
        });
    }

    private void engOverAllTeam(){

        txtSearchEng.clear();
        int readAllNum = readOverAllUsers.size();
        ObservableList<String> users = FXCollections.observableArrayList();

        for (int i = 0; i <readAllNum ; i++) {
            users.add(readOverAllUsers.get(i));
        }

        if (engNameListAll.getItems().size() == 0) {
            engNameListAll.getItems().addAll(users);
        }

        engNameListAll.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);

        lblSearcEng.setVisible(true);
        txtSearchEng.setVisible(true);

        FilteredList<String> filteredEng = new FilteredList((ObservableList) users, p -> true);

        txtSearchEng.textProperty().addListener((observable, oldValue, newValue) -> {
            filteredEng.setPredicate(string -> {

                engSkilLev.setVisible(false);
                engSkillName.setVisible(false);

                if (newValue == null || newValue.isEmpty()) {
                    return true;
                }

                String lowerCaseCustomerName = newValue.toLowerCase();

                if (string.toLowerCase().contains(lowerCaseCustomerName)) {
                    return true;
                }
                return false;
            });
        });

        engNameListAll.setItems(filteredEng);
        engNameListAll.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                engSkilLev.getItems().clear();
                engSkillName.getItems().clear();
                selected = "";
                selectedLevel = "";
                engSkilLev.setVisible(true);
                engSkillName.setVisible(false);

                selected = engNameListAll.getSelectionModel().getSelectedItem();
                userRef =0;
                skillLevelList.getItems().clear();
                levels.clear();
                setLevels();
                engSkilLev.getItems().addAll(levels);

                engSkilLev.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
                engSkilLev.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        engSkillName.getItems().clear();
                        selectedLevel = engSkilLev.getSelectionModel().getSelectedItem();
                        engSkillName.setVisible(true);

                        try {

                            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\SkillSet\\SkillsetProfiles.xls")));
                            HSSFSheet sheet = workbook.getSheetAt(0);
                            HSSFCell cellVal;

                            int cellnum = sheet.getRow(0).getLastCellNum();
                            int lastRow = sheet.getLastRowNum();
                            expertLevel = new ArrayList<>();
                            intLevel = new ArrayList<>();
                            basicLevel = new ArrayList<>();
                            noLevel = new ArrayList<>();

                            for (int i = 1; i <cellnum ; i++) {
                                String userNameColl = sheet.getRow(0).getCell(i).getStringCellValue();
                                if (userNameColl.equals(selected)){
                                    userRef = i;
                                }
                            }

                            for (int i = 1; i <lastRow ; i++) {

                                cellVal = sheet.getRow(i).getCell(userRef);

                                if (cellVal != null) {

                                    int cellValue = 0;

                                    if (cellVal.getCellType() == CellType.NUMERIC){
                                        cellValue = (int) cellVal.getNumericCellValue();
                                    }
                                    else{
                                        cellValue = Integer.parseInt(cellVal.getStringCellValue());
                                    }

                                    if (cellValue == 3) {
                                        expertLevel.add(sheet.getRow(i).getCell(0).getStringCellValue());
                                    }
                                    if (cellValue == 2) {
                                        intLevel.add(sheet.getRow(i).getCell(0).getStringCellValue());
                                    }
                                    if (cellValue == 1) {
                                        basicLevel.add(sheet.getRow(i).getCell(0).getStringCellValue());
                                    }
                                    if (cellValue == 0) {
                                        noLevel.add(sheet.getRow(i).getCell(0).getStringCellValue());
                                    }
                                }else {
                                    noLevel.add(sheet.getRow(i).getCell(0).getStringCellValue());
                                }
                            }

                        }catch (Exception e){
                            e.printStackTrace();
                        }

                        int expertSize = expertLevel.size();
                        int intermSize = intLevel.size();
                        int basicSize = basicLevel.size();
                        int noSize = noLevel.size();

                        ObservableList<String> skills = FXCollections.observableArrayList();

                        if (selectedLevel.equals("EXPERT")){

                            for (int i = 0; i <expertSize ; i++) {
                                skills.addAll(expertLevel.get(i));
                            }

                            engSkillName.getItems().addAll(skills);
                        }
                        if (selectedLevel.equals("INTERMEDIATE")){

                            for (int i = 0; i <intermSize ; i++) {
                                skills.addAll(intLevel.get(i));
                            }

                            engSkillName.getItems().addAll(skills);
                        }
                        if (selectedLevel.equals("BEGINNER")){

                            for (int i = 0; i <basicSize ; i++) {
                                skills.addAll(basicLevel.get(i));
                            }

                            engSkillName.getItems().addAll(skills);
                        }
                        if (selectedLevel.equals("NONE")){

                            for (int i = 0; i <noSize ; i++) {
                                skills.addAll(noLevel.get(i));
                            }

                            engSkillName.getItems().addAll(skills);
                        }

                    }
                });

            }
        });
    }

    private void skillMyTeam(){

        txtSearchSkill.clear();
        readSkills();

        int skillsAllSize = skillsAll.size();

        ObservableList<String> skills = FXCollections.observableArrayList();
        ObservableList<String> engins = FXCollections.observableArrayList();

        for (int i = 0; i <skillsAllSize ; i++) {
            skills.addAll(skillsAll.get(i));
        }

        if (skillNameList.getItems().size() == 0) {
            skillNameList.getItems().addAll(skills);
        }

        skillNameList.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);

        lblSearchSkill.setVisible(true);
        txtSearchSkill.setVisible(true);

        FilteredList<String> filteredSkill = new FilteredList((ObservableList) skills, p -> true);

        txtSearchSkill.textProperty().addListener((observable, oldValue, newValue) -> {
            filteredSkill.setPredicate(string -> {

                skillLevelList.setVisible(false);
                skillEngName.setVisible(false);
                btnExport.setVisible(false);

                if (newValue == null || newValue.isEmpty()) {
                    return true;
                }

                String lowerCaseCustomerName = newValue.toLowerCase();

                if (string.toLowerCase().contains(lowerCaseCustomerName)) {
                    return true;
                }
                return false;
            });
        });

        skillNameList.setItems(filteredSkill);

        skillNameList.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                selectedSkill = "";
                selectedSkillLevel ="";
                skillRef = 0;
                skillLevelList.setVisible(true);
                skillEngName.setVisible(false);
                btnExport.setVisible(false);

                selectedSkill = skillNameList.getSelectionModel().getSelectedItem();

                skillLevelList.getItems().clear();
                levels.clear();
                setLevels();
                skillLevelList.getItems().addAll(levels);

                skillLevelList.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        int compareNum = 0;

                        skillsExpert = new ArrayList<>();
                        skillsInterm = new ArrayList<>();
                        skillsBegin = new ArrayList<>();
                        skillsNone = new ArrayList<>();

                        skillEngName.setVisible(true);
                        btnExport.setVisible(true);

                        selectedSkillLevel = skillLevelList.getSelectionModel().getSelectedItem();

                        try {
                            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\SkillSet\\SkillsetProfiles.xls")));
                            HSSFSheet sheet = workbook.getSheetAt(0);
                            HSSFCell cellVal;
                            HSSFCell cellVal2;

                            int cellnum = sheet.getRow(0).getLastCellNum();
                            int lastRow = sheet.getLastRowNum();

                            for (int i = 1; i <lastRow ; i++) {
                                String skillNameColl = sheet.getRow(i).getCell(0).getStringCellValue();
                                if (skillNameColl.equals(selectedSkill)){
                                    skillRef = i;
                                }
                            }

                            skillsExpert.clear();
                            skillsInterm.clear();
                            skillsBegin.clear();
                            skillsNone.clear();

                            for (int i = 1; i < cellnum ; i++) {

                                cellVal = sheet.getRow(skillRef).getCell(i);

                                if (cellVal != null) {

                                    int cellValue = 0;

                                    if (cellVal.getCellType() == CellType.NUMERIC) {
                                        cellValue = (int) cellVal.getNumericCellValue();
                                    } else {
                                        cellValue = Integer.parseInt(cellVal.getStringCellValue());
                                    }

                                    if (cellValue == 3){

                                        cellVal2 = sheet.getRow(0).getCell(i);
                                        String engName = cellVal2.getStringCellValue();

                                        if (readUserList.contains(engName)) {
                                            skillsExpert.add(engName);
                                        }
                                    }
                                    if (cellValue == 2){

                                        cellVal2 = sheet.getRow(0).getCell(i);
                                        String engName = cellVal2.getStringCellValue();

                                        if (readUserList.contains(engName)) {
                                            skillsInterm.add(engName);
                                        }
                                    }
                                    if (cellValue == 1){

                                        cellVal2 = sheet.getRow(0).getCell(i);
                                        String engName = cellVal2.getStringCellValue();

                                        if (readUserList.contains(engName)) {
                                            skillsBegin.add(engName);
                                        }
                                    }
                                    if (cellValue == 0){

                                        cellVal2 = sheet.getRow(0).getCell(i);
                                        String engName = cellVal2.getStringCellValue();

                                        if (readUserList.contains(engName)) {
                                            skillsNone.add(engName);
                                        }
                                    }
                                }
                                else {
                                    cellVal2 = sheet.getRow(0).getCell(i);
                                    String engName = cellVal2.getStringCellValue();
                                    if (readUserList.contains(engName)) {
                                        skillsNone.add(engName);
                                    }
                                }
                            }

                            if (selectedSkillLevel.equals("EXPERT")){

                                engins.clear();
                                engins.addAll(skillsExpert);
                                skillEngName.getItems().clear();
                                skillEngName.getItems().addAll(engins);
                            }
                            if (selectedSkillLevel.equals("INTERMEDIATE")){

                                engins.clear();
                                engins.addAll(skillsInterm);
                                skillEngName.getItems().clear();
                                skillEngName.getItems().addAll(engins);
                            }
                            if (selectedSkillLevel.equals("BEGINNER")){

                                engins.clear();
                                engins.addAll(skillsBegin);
                                skillEngName.getItems().clear();
                                skillEngName.getItems().addAll(engins);
                            }
                            if (selectedSkillLevel.equals("NONE")){

                                engins.clear();
                                engins.addAll(skillsNone);
                                skillEngName.getItems().clear();
                                skillEngName.getItems().addAll(engins);
                            }

                        }catch (Exception e){
                            e.printStackTrace();
                        }

                        btnExport.setOnMouseClicked(new EventHandler<MouseEvent>() {
                            @Override
                            public void handle(MouseEvent event) {

                                try {

                                    FileChooser fileChooser = new FileChooser();
                                    FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("TXT Files (*.txt)", "*.txt");
                                    fileChooser.setInitialDirectory(new File(System.getProperty("user.home") + "\\Desktop"));
                                    fileChooser.setInitialFileName(selectedSkill + "_" + selectedSkillLevel + "_Level_Engineers");

                                    fileChooser.getExtensionFilters().add(extFilter);

                                    Stage primaryStage = new Stage();

                                    File file = fileChooser.showSaveDialog(primaryStage);

                                    FileWriter writer = new FileWriter(file);

                                    primaryStage.show();

                                    if (file != null) {

                                        int size = engins.size();
                                        for (int i = 0; i < size; i++) {
                                            String str = engins.get(i);
                                            writer.write(str);
                                            if (i < size - 1)
                                                writer.write("\n");
                                        }
                                        writer.close();
                                    }

                                    primaryStage.close();
                                } catch (Exception e) {
                                    e.printStackTrace();
                                }
                            }
                        });
                    }
                });

            }
        });
    }

    private void skillOverAllTeam(){

        txtSearchSkill.clear();
        readSkills();

        int skillsAllSize = skillsAll.size();

        ObservableList<String> skills = FXCollections.observableArrayList();
        ObservableList<String> engins = FXCollections.observableArrayList();

        for (int i = 0; i <skillsAllSize ; i++) {
            skills.addAll(skillsAll.get(i));
        }

        if (skillNameListAll.getItems().size() == 0) {
            skillNameListAll.getItems().addAll(skills);
        }

        skillNameListAll.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);

        lblSearchSkill.setVisible(true);
        txtSearchSkill.setVisible(true);

        FilteredList<String> filteredSkill = new FilteredList((ObservableList) skills, p -> true);

        txtSearchSkill.textProperty().addListener((observable, oldValue, newValue) -> {
            filteredSkill.setPredicate(string -> {

                if (newValue == null || newValue.isEmpty()) {
                    return true;
                }

                String lowerCaseCustomerName = newValue.toLowerCase();

                if (string.toLowerCase().contains(lowerCaseCustomerName)) {
                    return true;
                }
                return false;
            });
        });

        skillNameListAll.setItems(filteredSkill);

        skillNameListAll.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                selectedSkill = "";
                selectedSkillLevel ="";
                skillRef = 0;

                skillLevelList.setVisible(true);
                skillEngName.setVisible(false);
                btnExport.setVisible(false);

                selectedSkill = skillNameListAll.getSelectionModel().getSelectedItem();

                skillLevelList.getItems().clear();
                levels.clear();
                setLevels();
                skillLevelList.getItems().addAll(levels);

                skillLevelList.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        int compareNum = 0;

                        skillsExpert = new ArrayList<>();
                        skillsInterm = new ArrayList<>();
                        skillsBegin = new ArrayList<>();
                        skillsNone = new ArrayList<>();

                        skillEngName.setVisible(true);


                        selectedSkillLevel = skillLevelList.getSelectionModel().getSelectedItem();

                        try {
                            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\SkillSet\\SkillsetProfiles.xls")));
                            HSSFSheet sheet = workbook.getSheetAt(0);
                            HSSFCell cellVal;
                            HSSFCell cellVal2;

                            int cellnum = sheet.getRow(0).getLastCellNum();
                            int lastRow = sheet.getLastRowNum();

                            for (int i = 1; i <lastRow ; i++) {
                                String skillNameColl = sheet.getRow(i).getCell(0).getStringCellValue();
                                if (skillNameColl.equals(selectedSkill)){
                                    skillRef = i;
                                }
                            }

                            skillsExpert.clear();
                            skillsInterm.clear();
                            skillsBegin.clear();
                            skillsNone.clear();

                            for (int i = 1; i < cellnum ; i++) {

                                cellVal = sheet.getRow(skillRef).getCell(i);

                                if (cellVal != null) {

                                    int cellValue = 0;

                                    if (cellVal.getCellType() == CellType.NUMERIC) {
                                        cellValue = (int) cellVal.getNumericCellValue();
                                    } else {
                                        cellValue = Integer.parseInt(cellVal.getStringCellValue());
                                    }

                                    if (cellValue == 3){

                                        cellVal2 = sheet.getRow(0).getCell(i);
                                        String engName = cellVal2.getStringCellValue();

                                        skillsExpert.add(engName);
                                    }
                                    if (cellValue == 2){

                                        cellVal2 = sheet.getRow(0).getCell(i);
                                        String engName = cellVal2.getStringCellValue();

                                        skillsInterm.add(engName);
                                    }
                                    if (cellValue == 1){

                                        cellVal2 = sheet.getRow(0).getCell(i);
                                        String engName = cellVal2.getStringCellValue();

                                        skillsBegin.add(engName);
                                    }
                                    if (cellValue == 0){

                                        cellVal2 = sheet.getRow(0).getCell(i);
                                        String engName = cellVal2.getStringCellValue();

                                        skillsNone.add(engName);
                                    }
                                }
                                else {
                                    cellVal2 = sheet.getRow(0).getCell(i);
                                    String engName = cellVal2.getStringCellValue();
                                    skillsNone.add(engName);
                                }
                            }

                            if (selectedSkillLevel.equals("EXPERT")){

                                engins.clear();
                                engins.addAll(skillsExpert);
                                skillEngName.getItems().clear();
                                skillEngName.getItems().addAll(engins);
                            }
                            if (selectedSkillLevel.equals("INTERMEDIATE")){

                                engins.clear();
                                engins.addAll(skillsInterm);
                                skillEngName.getItems().clear();
                                skillEngName.getItems().addAll(engins);
                            }
                            if (selectedSkillLevel.equals("BEGINNER")){

                                engins.clear();
                                engins.addAll(skillsBegin);
                                skillEngName.getItems().clear();
                                skillEngName.getItems().addAll(engins);
                            }
                            if (selectedSkillLevel.equals("NONE")){

                                engins.clear();
                                engins.addAll(skillsNone);
                                skillEngName.getItems().clear();
                                skillEngName.getItems().addAll(engins);
                            }

                        }catch (Exception e){
                            e.printStackTrace();
                        }
                    }
                });

            }
        });

    }

    private void readSkills(){

        try {

            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\SkillSet\\SkillsetProfiles.xls")));
            HSSFSheet sheet = workbook.getSheetAt(0);
            HSSFCell cellVal;

            int skillRef = 0;
            int lastRow = sheet.getLastRowNum();

            skillsAll = new ArrayList<>();

            for (int i = 1; i <lastRow ; i++) {

                cellVal = sheet.getRow(i).getCell(skillRef);
                String skill = cellVal.getStringCellValue();
                skillsAll.add(skill);
            }

        } catch (Exception e){
            e.printStackTrace();
        }
    }

    private void readUsers(){

        File usersFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\SkillSet\\users.txt");

        if (usersFile.isFile()) {

            Scanner s = null;
            try {
                s = new Scanner(usersFile);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            readUserList = new ArrayList<String>();
            while (s.hasNextLine()) {
                readUserList.add(s.nextLine());
            }
            s.close();
        }
    }

    private void failSafeUsers(){

        int size = readUserList.size();

        safeUserList = new ArrayList<>();

        for (int i = 0; i <size ; i++) {

            if (readOverAllUsers.contains(readUserList.get(i))){
                safeUserList.add(readUserList.get(i));
            }
        }
    }

    private void readAllUsers(){

        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\SkillSet\\SkillsetProfiles.xls")));
            HSSFSheet sheet = workbook.getSheetAt(0);

            int cellnum = sheet.getRow(0).getLastCellNum();
            int lastRow = sheet.getLastRowNum();

            readOverAllUsers = new ArrayList<>();
            for (int i = 1; i <cellnum ; i++) {
                String userNameColl = sheet.getRow(0).getCell(i).getStringCellValue();
                readOverAllUsers.add(userNameColl);
            }

        }catch(Exception e){
            e.printStackTrace();
        }
    }

    private void setLevels(){

        ArrayList<String> level = new ArrayList<>();
        List lev = Arrays.asList("EXPERT", "INTERMEDIATE", "BEGINNER", "NONE");
        level.addAll(lev);
        levels.addAll(level);
    }

    @FXML
    void handleMouseClick(MouseEvent event) {

    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        rdEngMyTeam.setSelected(false);
        rdSkilMyTeam.setSelected(false);
        readAllUsers();
        readUsers();
        failSafeUsers();
    }
}