package home;

import de.jensd.fx.glyphs.fontawesome.FontAwesomeIcon;
import de.jensd.fx.glyphs.fontawesome.FontAwesomeIconView;
import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.collections.transformation.SortedList;
import javafx.concurrent.Worker;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.MenuItem;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.Image;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.Pane;
import javafx.scene.layout.VBox;
import javafx.scene.web.WebEngine;
import javafx.scene.web.WebView;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Duration;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.w3c.dom.Document;

import javax.naming.TimeLimitExceededException;
import java.awt.*;
import java.io.*;
import java.net.URL;
import java.sql.Array;
import java.sql.Time;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class Controller implements Initializable {

    @FXML
    private VBox vbox;
    @FXML
    private ProgressBar progressBar;
    @FXML
    private AnchorPane apnTableView;
    @FXML
    private AnchorPane apnSettings;
    @FXML
    private AnchorPane apnMyCases;
    @FXML
    private AnchorPane apnHome;
    @FXML
    private AnchorPane apnProduct;
    @FXML
    private AnchorPane apnCustomers;
    @FXML
    private AnchorPane apnProjects;
    @FXML
    private Pane pnProjectCaseTable;
    @FXML
    private AnchorPane apnBrowser;
    @FXML
    private Pane browserLoginPane;
    @FXML
    private Label lblStatus;
    @FXML
    private Label lblRefreshText;
    @FXML
    private TextField txUsers;
    @FXML
    private TextField txProducts;
    @FXML
    public TextField customerText;
    @FXML
    private TextField txQueues;
    @FXML
    private Button btnHome;
    @FXML
    private Button btnCases;
    @FXML
    private Button btnProducts;
    @FXML
    private Button btnProjects;
    @FXML
    private Button btnCustomers;
    @FXML
    private Button btnSurvey;
    @FXML
    private Button btnSettings;
    @FXML
    private Button btnLoadData;
    @FXML
    private Button btnLogin;
    @FXML
    private Button btnE1Cases;
    @FXML
    private Button btnMyE1Cases;
    @FXML
    private Button btnE2Cases;
    @FXML
    private Button btnMyE2Cases;
    @FXML
    private Button btnOutFollow;
    @FXML
    private Button btnMyOutFollow;
    @FXML
    private Button btnEscalated;
    @FXML
    private Button btnMyEscalated;
    @FXML
    private Button btnBCCases;
    @FXML
    private Button btnMyBCCases;
    @FXML
    private Button btnHotIssues;
    @FXML
    private Button btnMyHotIssues;
    @FXML
    private Button btnWOH;
    @FXML
    private Button btnMyWOH;
    @FXML
    private Button btnInactive;
    @FXML
    private Button btnMyInactive;
    @FXML
    private Button btnBCWIP;
    @FXML
    private Button btnMyBCWIP;
    @FXML
    private Button btnBCWac;
    @FXML
    private Button btnMyBCWac;
    @FXML
    private Button btnBCupdated;
    @FXML
    private Button btnMyBCupdated;
    @FXML
    private Button btnBCEngineering;
    @FXML
    private Button btnMyBCEngineering;
    @FXML
    private Button btnBCINACT;
    @FXML
    private Button btnMyBCINACT;
    @FXML
    private Button btnMJWIP;
    @FXML
    private Button btnMyMJWIP;
    @FXML
    private Button btnMJWac;
    @FXML
    private Button btnMyMJWac;
    @FXML
    private Button btnMJupdated;
    @FXML
    private Button btnMyMJupdated;
    @FXML
    private Button btnMJEngineering;
    @FXML
    private Button btnMyMJEngineering;
    @FXML
    private Button btnMJINACT;
    @FXML
    private Button btnMyMJINACT;
    @FXML
    private Button btnBCDue;
    @FXML
    private Button btnMyBCDue;
    @FXML
    private Button btnBCMissed;
    @FXML
    private Button btnMyBCMissed;
    @FXML
    private Button btnMJDue;
    @FXML
    private Button btnMyMJDue;
    @FXML
    private Button btnMJMissed;
    @FXML
    private Button btnMNMissed;
    @FXML
    private Button btnUpdateToday;
    @FXML
    private Button btnUpdateMissed;
    @FXML
    private Button btnUpdateNull;
    @FXML
    private Button btnMyUpdateToday;
    @FXML
    private Button btnMyUpdateMissed;
    @FXML
    private Button btnMyUpdateNull;
    @FXML
    private Button btnMyMJMissed;
    @FXML
    private Button btnMyQueue;
    @FXML
    private Button btnPSQueue;
    @FXML
    private Button btnTSQueue;
    @FXML
    private Button btnE1Prod;
    @FXML
    private Button btnE2Prod;
    @FXML
    private Button btnOutFollowProd;
    @FXML
    private Button btnEscalatedProd;
    @FXML
    private Button btnBCProd;
    @FXML
    private Button btnHotIssuesProd;
    @FXML
    private Button btnWOHProd;
    @FXML
    private Button btnInactiveProd;
    @FXML
    private Button btnBCWIPProd;
    @FXML
    private Button btnBCWacProd;
    @FXML
    private Button btnBCupdatedProd;
    @FXML
    private Button btnBCEngineeringProd;
    @FXML
    private Button btnBCINACTProd;
    @FXML
    private Button btnMJupdatedProd;
    @FXML
    private Button btnMJWacProd;
    @FXML
    private Button btnMJWIPProd;
    @FXML
    private Button btnMJINACTProd;
    @FXML
    private Button btnMJEngineeringProd;
    @FXML
    private Button btnMJDueProd;
    @FXML
    private Button btnBCMissedProd;
    @FXML
    private Button btnBCDueProd;
    @FXML
    private Button btnMJMissedProd;
    @FXML
    private Button btnPSQueueProd;
    @FXML
    private Button btnTSQueueProd;
    @FXML
    private Button btnAccountClear;
    @FXML
    private Button btnCustomerLoad;
    @FXML
    private Button btnCustomerCritical;
    @FXML
    private Button btnCustomerE2;
    @FXML
    private Button btnCustomerOutFollow;
    @FXML
    private Button btnCustomerEscalated;
    @FXML
    private Button btnCustomerHotIssues;
    @FXML
    private Button btnCustomerBC;
    @FXML
    private Button btnCustomerActWOH;
    @FXML
    private FontAwesomeIconView btnBack;
    @FXML
    private FontAwesomeIconView btnInfo;
    @FXML
    private FontAwesomeIconView btnToExcel;
    @FXML
    private TableView<CaseTableView> tableCases;
    @FXML
    private TableView<CaseTableView> tableCustomers;
    @FXML
    private TableColumn<CaseTableView, String> NumberCol;
    @FXML
    private TableColumn<CaseTableView, String> StatusCol;
    @FXML
    private TableColumn<CaseTableView, String> SeverityCol;
    @FXML
    private TableColumn<CaseTableView, String> ResponsibleCol;
    @FXML
    private TableColumn<CaseTableView, String> OwnerCol;
    @FXML
    private TableColumn<CaseTableView, String> EscalatedByCol;
    @FXML
    private TableColumn<CaseTableView, String> HotListCol;
    @FXML
    private TableColumn<CaseTableView, String> AgeCol;
    @FXML
    private TableColumn<CaseTableView, String> ProductCol;
    @FXML
    private TableColumn<CaseTableView, String> AccountCol;
    @FXML
    private TableColumn<CaseTableView, String> SubjectCol;
    @FXML
    private TableColumn<CaseTableView, String> OutFollowCol;
    @FXML
    private TableColumn<CaseTableView, String> SupportTypeCol;
    @FXML
    private TableColumn<CaseTableView, LocalDate> NextUpdateCol;
    @FXML
    private TableColumn<CaseTableView, String> DateTimeOpenedCol;
    @FXML
    private TableColumn<CaseTableView, String> RegionCol;
    @FXML
    private TableColumn<CaseTableView, String> SecurityCol;
    @FXML
    private TableColumn<CaseTableView, String> NumberColCust;
    @FXML
    private TableColumn<CaseTableView, String> StatusColCust;
    @FXML
    private TableColumn<CaseTableView, String> SeverityColCust;
    @FXML
    private TableColumn<CaseTableView, String> ResponsibleColCust;
    @FXML
    private TableColumn<CaseTableView, String> OwnerColCust;
    @FXML
    private TableColumn<CaseTableView, String> EscalatedByColCust;
    @FXML
    private TableColumn<CaseTableView, String> HotListColCust;
    @FXML
    private TableColumn<CaseTableView, String> AgeColCust;
    @FXML
    private TableColumn<CaseTableView, String> ProductColCust;
    @FXML
    private TableColumn<CaseTableView, String> AccountColCust;
    @FXML
    private TableColumn<CaseTableView, String> SubjectColCust;
    @FXML
    private TableColumn<CaseTableView, String> OutFollowColCust;
    @FXML
    private TableColumn<CaseTableView, String> SupportTypeColCust;
    @FXML
    private TableColumn<CaseTableView, LocalDate> NextUpdateColCust;
    @FXML
    private TableColumn<CaseTableView, String> DateTimeOpenedColCust;
    @FXML
    private TableColumn<CaseTableView, String> RegionColCust;
    @FXML
    private TableColumn<CaseTableView, String> SecurityColCust;
    @FXML
    private TableView<AccountTableView> tableAccounts;
    @FXML
    private TableColumn<AccountTableView, String> customerCol;
    @FXML
    private Pane pnAccountSelect;
    @FXML
    private Button btnFilterAccountAdd;
    @FXML
    private Button btnFilterAccountClear;
    @FXML
    private Button btnFilterAccountUpdate;
    @FXML
    private Button btnFilterAccountClose;
    @FXML
    private TextField txtFilterAccounts;
    @FXML
    private TableView<AccountTableView> tableAccountsSelected;
    @FXML
    private TableColumn<AccountTableView, String> customerSelectedCol;
    @FXML
    private TableView<UserTableView> tableUsers;
    @FXML
    private TableView<UserTableView> tableUsersSelected;
    @FXML
    private TableColumn<UserTableView, String> userCol;
    @FXML
    private TableColumn<UserTableView, String> userSelectedCol;
    @FXML
    private Button btnUsersUpdate;
    @FXML
    private Button btnUserSelectClose;
    @FXML
    private Pane pnUsersSelect;
    @FXML
    private Button btnUsersClear;
    @FXML
    private Button btnProductsClear;
    @FXML
    private Button btnQueueClear;
    @FXML
    private TextField txtUserSelect;
    @FXML
    private TextField txtProductSelect;
    @FXML
    private TableView<ProductTableView> tableProducts;
    @FXML
    private TableView<ProductTableView> tableProductsSelected;
    @FXML
    private TableColumn<ProductTableView, String> productCol;
    @FXML
    private TableColumn<ProductTableView, String> productColSelected;
    @FXML
    private Button btnProductUpdate;
    @FXML
    private Button btnProductSelectClose;
    @FXML
    private Pane pnProductSelect;
    @FXML
    private TextField txtQueueSelect;
    @FXML
    private TableView<QueueTableView> tableQueue;
    @FXML
    private TableView<QueueTableView> tableQueueSelected;
    @FXML
    private TableColumn<QueueTableView, String> queueCol;
    @FXML
    private TableColumn<QueueTableView, String> queueColSelected;
    @FXML
    private Button btnQueueUpdate;
    @FXML
    private Button btnQueueSelectClose;
    @FXML
    private Pane pnQueueSelect;
    @FXML
    private Button btnClearAll;
    @FXML
    private Button btnSaveDefault;
    @FXML
    private Button btnLoadDefault;
    @FXML
    private Button btnUserSelectClear;
    @FXML
    private Button btnProductSelectClear;
    @FXML
    private Button btnQueueSelectClear;
    Timeline time = new Timeline();


    WebView browserLogin = new WebView();
    ArrayList<String> settingsUsers = new ArrayList<>();
    ArrayList<String> settingsQueue = new ArrayList<>();
    ArrayList<String> settingsProducts = new ArrayList<>();
    ArrayList<String> filteredAccounts = new ArrayList<String>();
    ArrayList<String> usersFiltered = new ArrayList<String>();
    ArrayList<String> productsFiltered = new ArrayList<String>();
    ArrayList<String> queuesFiltered = new ArrayList<String>();
    ArrayList<String> queueArray = new ArrayList<>();

    ContextMenu menu = new ContextMenu();
    MenuItem openCaseSFDC = new MenuItem("Search case in SalesForce...");


    //Case Ref Cells
    int caseAccountRef = 0;
    int caseNumCellRef = 0;
    int caseSupTypeRefCell = 0;
    int caseStatRefCell = 0;
    int caseSevRefCell = 0;
    int caseRespRefCell = 0;
    int caseOwnerRefCell = 0;
    int caseEscalatedRefCell = 0;
    int caseHotListRefCell = 0;
    int caseOutFolRefCell = 0;
    int caseAgeRefCell = 0;
    int caseMainQuestionRefCell = 0;
    int caseSubQuestionRefCell = 0;
    int caseAnswerRefCell = 0;
    int caseSurveyTypeRef;
    int mycaseNumCellRef = 0;
    int mycaseSupTypeRefCell = 0;
    int mycaseStatRefCell = 0;
    int mycaseSevRefCell = 0;
    int mycaseRespRefCell = 0;
    int mycaseOwnerRefCell = 0;
    int mycaseEscalatedRefCell = 0;
    int mycaseHotListRefCell = 0;
    int mycaseOutFolRefCell = 0;
    int mycaseAgeRefCell = 0;
    int mycaseUpdateCell = 0;
    int caseCellRef = 0;
    int caseCellRef2 = 0;
    int myCaseCellRef1 = 0;
    int caseNextUpdateDateRef = 0;
    int caseProductRef = 0;
    int customerE1 = 0;
    int customerE2 = 0;
    int customerOutFol = 0;
    int customerHot = 0;
    int customerEsc = 0;
    int customerBC = 0;
    int customerWoh = 0;

    //Overview Page # variables
    int e1Cases = 0;
    int e2Cases = 0;
    int outFollow = 0;
    int queueTS = 0;
    int queuePS = 0;
    int updateToday = 0;
    int updateMissed = 0;
    int updateNull = 0;
    int hotlist = 0;
    int escCase = 0;
    int bcCases = 0;
    int inactiveCases = 0;
    int wohCases = 0;
    int bcDue = 0;
    int misBCdue = 0;
    int custActBC = 0;
    int custRpdBC = 0;
    int BCds = 0;
    int BCpc = 0;
    int BCwip = 0;
    int dueMJday = 0;
    int misMJdue = 0;
    int misMNdue = 0;
    int custActMJ = 0;
    int custRpdMJ = 0;
    int MJds = 0;
    int MJpc = 0;
    int MJwip = 0;

    //My Page # variables
    int myHotList = 0;
    int myOutFollow = 0;
    int myEscCases = 0;
    int myBCCases = 0;
    int myInactiveCases = 0;
    int myBCDueCases = 0;
    int myBCMissedCases = 0;
    int myBCDSCases = 0;
    int myBCInactiveCases = 0;
    int myBCWIP = 0;
    int myMJDueCases = 0;
    int myMJMissedCases = 0;
    int myMJUpdated = 0;
    int myMJDSCases = 0;
    int myMJWIP = 0;
    int myQueuedCases = 0;
    int myE1Case = 0;
    int myE2Cases = 0;
    int myBCupdated = 0;
    int myBCWac = 0;
    int myMJWAC = 0;
    int myMJInactiveCases = 0;
    int myWOHCases = 0;
    int myUpdateToday = 0;
    int myUpdateMissed = 0;
    int myUpdateNull = 0;

    //Product Page # variables

    int prodHotList = 0;
    int prodOutFollow = 0;
    int prodEscCases = 0;
    int prodBCCases = 0;
    int prodInactiveCases = 0;
    int prodBCDueCases = 0;
    int prodBCMissedCases = 0;
    int prodBCDSCases = 0;
    int prodBCInactiveCases = 0;
    int prodBCWIP = 0;
    int prodMJDueCases = 0;
    int prodMJMissedCases = 0;
    int prodMJUpdated = 0;
    int prodMJDSCases = 0;
    int prodMJWIP = 0;
    int prodQueuedCases = 0;
    int prodE1Case = 0;
    int prodE2Cases = 0;
    int prodBCupdated = 0;
    int prodBCWac = 0;
    int prodMJWAC = 0;
    int prodMJInactiveCases = 0;
    int prodWOHCases = 0;
    int prodUpdateToday = 0;
    int prodUpdateMissed = 0;
    int prodUpdateNull = 0;
    int prodQueuePS = 0;
    int prodQueueTS = 0;

    @FXML
    private void handleClicks(ActionEvent event) throws IOException, InvalidFormatException {

        if (event.getSource() == btnHome) {
            lblStatus.setText("GENERAL OVERVIEW");
            btnToExcel.setVisible(false);
            apnHome.toFront();
            overviewPage();
        }

        if (event.getSource() == btnCases) {
            lblStatus.setText("MY CASES");
            btnToExcel.setVisible(false);
            myCasesPage();
            apnMyCases.toFront();
        }

        if (event.getSource() == btnProducts) {
            lblStatus.setText("PRODUCT VIEW");
            btnToExcel.setVisible(false);
            myProductsPage();
            apnProduct.toFront();
        }

        if (event.getSource() == btnProjects) {

        }

        if (event.getSource() == btnCustomers) {
            lblStatus.setText("CUSTOMER VIEW");
            tableCustomers.setVisible(false);
            apnCustomers.toFront();
            btnCustomerLoad.setVisible(false);
            pnAccountSelect.setVisible(false);
            accountArray();
        }

        if (event.getSource() == btnCustomerLoad) {
            if (!customerText.getText().isEmpty()) {
                customerViewPage();
            }
            tableCustomers.setVisible(false);
        }

        if (event.getSource() == btnSurvey) {
            /*parseSurveyData();
            readSurveyData();

            FontAwesomeIconView icon = new FontAwesomeIconView(FontAwesomeIcon.APPLE);
            icon.setSize("22");
            icon.getStyleClass().add("green-icon");
            btnSurvey.setGraphic(icon);*/
        }

        if (event.getSource() == btnSettings) {
            lblStatus.setText("SETTINGS");
            btnToExcel.setVisible(false);
            apnSettings.toFront();

        }

        if (event.getSource() == btnLogin) {

            //Connect to OKTA SSO
            connectOkta();
        }

        if (event.getSource() == btnLoadData) {

            //Download the related reports to work on them
            downloadCSV();
            parseUserData();
            apnHome.toFront();

        }

        if (event.getSource() == btnE1Cases) {
            if (e1Cases != 0) {
                lblStatus.setText("E1 - OUTAGE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter1 = "Critical";
                initTableView(tableCases);
                oneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
            }
            if (e1Cases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnE2Cases) {
            if (e2Cases != 0) {
                lblStatus.setText("E2 CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter1 = "E2";
                initTableView(tableCases);
                oneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
            }
            if (e2Cases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnOutFollow) {
            if (outFollow != 0) {
                lblStatus.setText("OUTAGE FOLLOW-UP CASES");
                tableCases.getItems().clear();
                String columnSelect = "Outage Follow-Up";
                String filter1 = "1";
                initTableView(tableCases);
                oneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
            }
            if (outFollow == 0) {
                alertUser();
            }

        }

        if (event.getSource() == btnEscalated) {
            if (escCase != 0) {
                lblStatus.setText("ESCALATED CASES");
                tableCases.getItems().clear();
                String columnSelect = "Escalated By";
                String filter1 = "NotSet";
                initTableView(tableCases);
                oneFilterTableView(columnSelect, filter1, tableCases, apnTableView, false);
            }
            if (escCase == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCCases) {
            if (bcCases != 0) {
                lblStatus.setText("BUSINESS CRITICAL CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter1 = "Business Critical";
                initTableView(tableCases);
                oneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
            }
            if (bcCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnHotIssues) {
            if (hotlist != 0) {
                lblStatus.setText("HOT ISSUES");
                tableCases.getItems().clear();
                String columnSelect = "Support Hotlist Level";
                String filter1 = "NotSet";
                initTableView(tableCases);
                oneFilterTableView(columnSelect, filter1, tableCases, apnTableView, false);
            }
            if (hotlist == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnWOH) {
            if (wohCases != 0) {
                lblStatus.setText("WORK ON HAND CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                overviewWOHView(true);
            }
            if (wohCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnInactive) {
            if (inactiveCases != 0) {
                lblStatus.setText("INACTIVE(PENDING CLOSURE) CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                overviewWOHView(false);
            }
            if (inactiveCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCWIP) {
            if (BCwip != 0) {
                lblStatus.setText("BUSINESS CRITICAL WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                overviewWIPCaseTableView(columnSelect, filter, tableCases, apnTableView);
            }
            if (BCwip == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCWac) {
            if (custActBC != 0) {
                lblStatus.setText("BUSINESS CRITICAL CASES PENDING INFORMATION FROM CUSTOMER");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Business Critical";
                String filter2 = "Customer action";
                initTableView(tableCases);
                twoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
            }
            if (custActBC == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCupdated) {
            if (custRpdBC != 0) {
                lblStatus.setText("BUSINESS CRITICAL CASES PENDING OWNER ACTION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Business Critical";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
            }
            if (custRpdBC == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCEngineering) {
            if (BCds != 0) {
                lblStatus.setText("BUSINESS CRITICAL CASES WITH DESIGN");
                tableCases.getItems().clear();
                String columSelect = "Severity";
                String filter1 = "Business Critical";
                initTableView(tableCases);
                overviewEngineeringTableView(columSelect, filter1, tableCases, apnTableView);
            }
            if (BCds == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCINACT) {
            if (BCpc != 0) {
                lblStatus.setText("BUSINESS CRITICAL CASES PENDING CLOSURE");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String filter1 = "Business Critical";
                initTableView(tableCases);
                overViewInactiveTable(columnSelect1, filter1, tableCases, apnTableView);
            }
            if (BCpc == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJWIP) {
            if (MJwip != 0) {
                lblStatus.setText("MAJOR WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                lblStatus.setText("BUSINESS CRITICAL WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                overviewWIPCaseTableView(columnSelect, filter, tableCases, apnTableView);
            }
            if (MJwip == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJWac) {
            if (custActMJ != 0) {
                lblStatus.setText("MAJOR CASES PENDING INFORMATION FROM CUSTOMER");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Major";
                String filter2 = "Customer action";
                initTableView(tableCases);
                twoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
            }
            if (custActMJ == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJupdated) {
            if (custRpdMJ != 0) {
                lblStatus.setText("MAJOR CASES PENDING OWNER ACTION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Major";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
            }
            if (custRpdMJ == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJEngineering) {
            if (MJds != 0) {
                lblStatus.setText("MAJOR CASES WITH DESIGN");
                tableCases.getItems().clear();
                tableCases.getItems().clear();
                String columSelect = "Severity";
                String filter1 = "Major";
                initTableView(tableCases);
                overviewEngineeringTableView(columSelect, filter1, tableCases, apnTableView);
            }
            if (MJds == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJINACT) {
            if (MJpc != 0) {
                lblStatus.setText("MAJOR CASES PENDING CLOSURE");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String filter1 = "Major";
                initTableView(tableCases);
                overViewInactiveTable(columnSelect1, filter1, tableCases, apnTableView);
            }
            if (MJpc == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCDue) {
            if (bcDue != 0) {
                lblStatus.setText("BUSINESS CRITICAL DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, true);
            }
            if (bcDue == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCMissed) {
            if (misBCdue != 0) {
                lblStatus.setText("BUSINESS CRITICAL CASES MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, false);
            }
            if (misBCdue == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJDue) {
            if (dueMJday != 0) {
                lblStatus.setText("MAJOR DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, true);
            }
            if (dueMJday == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJMissed) {
            if (misMJdue != 0) {
                lblStatus.setText("MAJOR CASES MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, false);
            }
            if (misMJdue == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMNMissed) {
            if (misMNdue != 0) {
                lblStatus.setText("MINOR CASES MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 180, false);
            }
            if (misMNdue == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnTSQueue) {
            if (queueTS != 0) {
                lblStatus.setText("CASES IN RTS QUEUE");
                tableCases.getItems().clear();
                String columnselect = "Case Owner";
                String filter = "TS";
                initTableView(tableCases);
                overviewQueueView(columnselect, filter, tableCases, apnTableView, "GPS QUEUE");
            }
            if (queueTS == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnPSQueue) {
            if (queuePS != 0) {
                lblStatus.setText("CASES IN GPS QUEUE");
                tableCases.getItems().clear();
                String e2TableSelect = "Case Owner";
                String e2TableSelect2 = "PS";
                initTableView(tableCases);
                overviewQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView, "RTS QUEUE");
            }
            if (queuePS == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnUpdateToday) {
            if (updateToday != 0) {
                lblStatus.setText("NEXT CASE UPDATE TODAY LIST");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                caseUpdateTableView(columnSelect, tableCases, apnTableView, true, true);
            }
            if (updateToday == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnUpdateMissed) {
            if (updateMissed != 0) {
                lblStatus.setText("NEXT CASE UPDATE MISSED LIST");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                caseUpdateTableView(columnSelect, tableCases, apnTableView, false, true);
            }
            if (updateMissed == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnUpdateNull) {
            if (updateNull != 0) {
                lblStatus.setText("NEXT CASE UPDATE FIELD NOT SET CASE LIST");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                caseUpdateTableView(columnSelect, tableCases, apnTableView, false, false);
            }
            if (updateNull == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyE1Cases) {
            if (myE1Case != 0) {
                lblStatus.setText("MY E1 CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Critical";
                initTableView(tableCases);
                oneFilterMyTableView(columnSelect, filter, tableCases, apnTableView, true);
            }
            if (myE1Case == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyE2Cases) {
            if (myE2Cases != 0) {
                lblStatus.setText("MY E2 CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "E2";
                initTableView(tableCases);
                oneFilterMyTableView(columnSelect, filter, tableCases, apnTableView, true);
            }
            if (myE2Cases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyOutFollow) {
            if (myOutFollow != 0) {

                lblStatus.setText("MY OUTAGE FOLLOW-UP CASES");
                tableCases.getItems().clear();
                String columnSelect = "Outage Follow-Up";
                String filter = "1";
                initTableView(tableCases);
                oneFilterMyTableView(columnSelect, filter, tableCases, apnTableView, true);
            }
            if (myOutFollow == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyEscalated) {
            if (myEscCases != 0) {
                lblStatus.setText("MY ESCALATED CASES");
                tableCases.getItems().clear();
                String columnSelect = "Escalated By";
                String filter = "NotSet";
                initTableView(tableCases);
                oneFilterMyTableView(columnSelect, filter, tableCases, apnTableView, false);
            }
            if (myEscCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyBCCases) {
            if (myBCCases != 0) {
                lblStatus.setText("MY BUSINESS CRITICAL CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                oneFilterMyTableView(columnSelect, filter, tableCases, apnTableView, true);
            }
            if (myBCCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyHotIssues) {
            if (myHotList != 0) {
                lblStatus.setText("MY HOT ISSUES");
                tableCases.getItems().clear();
                String columnSelect = "Support Hotlist Level";
                String filter = "NotSet";
                initTableView(tableCases);
                oneFilterMyTableView(columnSelect, filter, tableCases, apnTableView, false);
            }
            if (myHotList == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyWOH) {
            if (myWOHCases != 0) {
                lblStatus.setText("MY WORK ON HAND CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                myWOHTableView(tableCases, apnTableView, true);
            }
            if (myWOHCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyInactive) {
            if (myInactiveCases != 0) {
                lblStatus.setText("MY INACTIVE (PENDING CLOSURE) CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                myWOHTableView(tableCases, apnTableView, false);
            }
            if (myInactiveCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyBCWIP) {

            if (myBCWIP != 0) {

                lblStatus.setText("MY BUSINESS CRITICAL WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                String columFilter = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                overviewMyWIPCaseTableView(columFilter, filter, tableCases, apnTableView);
            }
            if (myBCWIP == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyBCWac) {

            if (myBCWac != 0) {
                lblStatus.setText("MY BUSINESS CRITICAL CASES PENDING INFORMATION FROM CUSTOMER");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Business Critical";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer action";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (myBCWac == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyBCupdated) {
            if (myBCupdated != 0) {

                lblStatus.setText("MY BUSINESS CRITICAL CASES PENDING OWNER ACTION");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Business Critical";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (myBCupdated == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyBCEngineering) {

            if (myBCDSCases != 0) {
                lblStatus.setText("MY BUSINESS CRITICAL CASES WITH DESIGN");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Business Critical";
                String columSelect2 = "Status";
                String filter2 = "Develop Solution";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (myBCDSCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyBCINACT) {

            if (myBCInactiveCases != 0) {

                lblStatus.setText("MY BUSINESS CRITICAL INACTIVE (PENDING CLOSURE) CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                inactiveCasesMyTableView(columnSelect, filter, tableCases, apnTableView);
            }
            if (myBCInactiveCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyMJWIP) {

            if (myMJWIP != 0) {

                lblStatus.setText("MY MAJOR WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                String columFilter = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                overviewMyWIPCaseTableView(columFilter, filter, tableCases, apnTableView);
            }
            if (myMJWIP == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyMJWac) {

            if (myMJWAC != 0) {
                lblStatus.setText("MY MAJOR CASES PENDING INFORMATION FROM CUSTOMER");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Major";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer action";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (myMJWAC == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyMJupdated) {

            if (myMJUpdated != 0) {
                String columSelect1 = "Severity";
                String filter1 = "Major";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (myMJUpdated == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyMJEngineering) {

            if (myMJDSCases != 0) {

                lblStatus.setText("MY MAJOR CASES WITH DESIGN");
                tableCases.getItems().clear();
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Major";
                String columSelect2 = "Status";
                String filter2 = "Develop Solution";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }

            if (myMJDSCases == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnMyMJINACT) {

            if (myMJInactiveCases != 0) {

                lblStatus.setText("MY MAJOR INACTIVE (PENDING CLOSURE) CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                inactiveCasesMyTableView(columnSelect, filter, tableCases, apnTableView);
            }
            if (myMJInactiveCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyBCDue) {

            if (myBCDueCases != 0) {
                lblStatus.setText("MY BUSINESS CRITICAL DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                myDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, true);
            }
            if (myBCDueCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyBCMissed) {

            if (myBCMissedCases != 0) {
                lblStatus.setText("MY BUSINESS CRITICAL CASED MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                myDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, false);
            }

            if (myBCMissedCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyMJDue) {

            if (myMJDueCases != 0) {
                lblStatus.setText("MY MAJOR DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                myDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, true);
            }

            if (myMJDueCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyMJMissed) {

            if (myMJMissedCases != 0) {
                lblStatus.setText("MY MAJOR CASED MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                myDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, false);
            }

            if (myMJMissedCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyQueue) {

            if (myQueuedCases != 0) {
                lblStatus.setText("CASES IN MY QUEUE(S)");
                tableCases.getItems().clear();
                String columnSelect = "Case Owner";
                initTableView(tableCases);
                createMyQueueCaseView(columnSelect, tableCases, apnTableView);
            }
            if (myQueuedCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyUpdateToday) {

            if (myUpdateToday != 0) {

                lblStatus.setText("NEXT CASE UPDATE TODAY LIST");
                tableCases.getItems().clear();
                String caseTableSelect = "Next Case Update";
                initTableView(tableCases);
                mycaseUpdateTableView(caseTableSelect, tableCases, apnTableView, true, true);
            }

            if (myUpdateToday == 0) {
                alertUser();
            }

        }

        if (event.getSource() == btnMyUpdateMissed) {

            if (myUpdateMissed != 0) {

                lblStatus.setText("NEXT CASE UPDATE MISSED LIST");
                tableCases.getItems().clear();
                String caseTableSelect = "Next Case Update";
                initTableView(tableCases);
                mycaseUpdateTableView(caseTableSelect, tableCases, apnTableView, false, true);
            }
            if (myUpdateMissed == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMyUpdateNull) {
            if (myUpdateNull != 0) {

                lblStatus.setText("NEXT CASE UPDATE FIELD NOT SET CASE LIST");
                tableCases.getItems().clear();
                String caseTableSelect = "Next Case Update";
                initTableView(tableCases);
                mycaseUpdateTableView(caseTableSelect, tableCases, apnTableView, false, false);
            }
            if (myUpdateNull == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnE1Prod) {
            if (prodE1Case != 0) {
                lblStatus.setText("PRODUCT VIEW - OUTAGE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Critical";
                initTableView(tableCases);
                productOneFilterView(columnSelect, filter, tableCases, apnTableView, true);
            }

            if (prodE1Case == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnE2Prod) {
            if (prodE2Cases != 0) {
                lblStatus.setText("PRODUCT VIEW - E2 CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "E2";
                initTableView(tableCases);
                productOneFilterView(columnSelect, filter, tableCases, apnTableView, true);
            }
            if (prodE2Cases == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnOutFollowProd) {
            if (prodOutFollow != 0) {
                lblStatus.setText("PRODUCT VIEW - OUTAGE FOLLOW-UP CASES");
                tableCases.getItems().clear();
                String columnSelect = "Outage Follow-Up";
                String filter = "1";
                initTableView(tableCases);
                productOneFilterView(columnSelect, filter, tableCases, apnTableView, true);
            }
            if (prodOutFollow == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnEscalatedProd) {
            if (prodEscCases != 0) {
                lblStatus.setText("PRODUCT VIEW - ESCALATED CASES");
                tableCases.getItems().clear();
                String columnSelect = "Escalated By";
                String filter = "NotSet";
                initTableView(tableCases);
                productOneFilterView(columnSelect, filter, tableCases, apnTableView, false);
            }
            if (prodEscCases == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnBCProd) {
            if (prodBCCases != 0) {
                lblStatus.setText("PRODUCT VIEW - BUSINESS CRITICAL CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                productOneFilterView(columnSelect, filter, tableCases, apnTableView, true);
            }
            if (prodBCCases == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnHotIssuesProd) {
            if (prodHotList != 0) {
                lblStatus.setText("PRODUCT VIEW - HOT ISSUES");
                tableCases.getItems().clear();
                String columnSelect = "Support Hotlist Level";
                String filter = "NotSet";
                initTableView(tableCases);
                productOneFilterView(columnSelect, filter, tableCases, apnTableView, false);
            }
            if (prodHotList == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnWOHProd) {
            if (prodWOHCases != 0) {
                lblStatus.setText("PRODUCT VIEW - ACTIVE WORK ON HAND");
                tableCases.getItems().clear();
                initTableView(tableCases);
                prodWOHTable(tableCases, apnTableView, true);
            }
            if (prodWOHCases == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnInactiveProd) {
            if (prodInactiveCases != 0) {
                lblStatus.setText("PRODUCT VIEW - INACTIVE WORK ON HAND");
                tableCases.getItems().clear();
                initTableView(tableCases);
                prodWOHTable(tableCases, apnTableView, false);
            }
            if (prodInactiveCases == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnBCWIPProd) {
            if (prodBCWIP != 0) {
                lblStatus.setText("PRODUCT VIEW - BC WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                productWIPCaseView(columnSelect, filter, tableCases, apnTableView);
            }
            if (prodBCWIP == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCWacProd) {
            if (prodBCWac != 0) {
                lblStatus.setText("PRODUCT VIEW - BC CASES PENDING CUSTOMER ACTION");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Business Critical";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer action";
                initTableView(tableCases);
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (prodBCWac == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCupdatedProd) {
            if (prodBCupdated != 0) {
                lblStatus.setText("PRODUCT VIEW - BC CASES CUSTOMER PROVIDED UPDATE");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Business Critical";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (prodBCupdated == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnBCEngineeringProd) {
            if (prodBCDSCases != 0) {
                lblStatus.setText("PRODUCT VIEW - BC CASES WITH DESIGN");
                tableCases.getItems().clear();
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Business Critical";
                String columSelect2 = "Status";
                String filter2 = "Develop Solution";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (prodBCDSCases == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnBCINACTProd) {
            if (prodBCInactiveCases != 0) {

                lblStatus.setText("PRODUCT VIEW - BC INACTIVE (PC & FA) CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                inactiveCasesProductTableView(columnSelect, filter, tableCases, apnTableView);
            }
            if (prodBCInactiveCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJWIPProd) {
            if (prodMJWIP != 0) {
                lblStatus.setText("MAJOR WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                String columnSelect = "Severity";
                String filter = "Major";
                productWIPCaseView(columnSelect, filter, tableCases, apnTableView);
            }
            if (prodMJWIP == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJWacProd) {
            if (prodMJWAC != 0) {
                lblStatus.setText("MAJOR CASES PENDING CUSTOMER ACTION");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Major";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer action";
                initTableView(tableCases);
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (prodMJWAC == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJupdatedProd) {
            if (prodMJUpdated != 0) {
                lblStatus.setText("MAJOR CASES CUSTOMER PROVIDED UPDATE");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Major";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (prodMJUpdated == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJEngineeringProd) {
            if (prodMJDSCases != 0) {
                lblStatus.setText("PRODUCT VIEW - MAJOR CASES WITH DESIGN");
                tableCases.getItems().clear();
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Major";
                String columSelect2 = "Status";
                String filter2 = "Develop Solution";
                initTableView(tableCases);
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
            }
            if (prodMJDSCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJINACTProd) {
            if (prodMJInactiveCases != 0) {

                lblStatus.setText("PRODUCT VIEW - MAJOR INACTIVE (PC & FA) CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                inactiveCasesProductTableView(columnSelect, filter, tableCases, apnTableView);
            }
            if (prodMJInactiveCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCDueProd) {

            if (prodBCDueCases != 0) {
                lblStatus.setText("PRODUCT VIEW - BUSINESS CRITICAL DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                productDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, true);
            }
            if (prodBCDueCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnBCMissedProd) {

            if (prodBCMissedCases != 0) {
                lblStatus.setText("PRODUCT VIEW - BUSINESS CRITICAL CASES MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                productDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, false);
            }
            if (prodBCMissedCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJDueProd) {

            if (prodMJDueCases != 0) {
                lblStatus.setText("PRODUCT VIEW - MAJOR DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                productDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, true);
            }
            if (prodMJDueCases == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnMJMissedProd) {

            if (prodMJMissedCases != 0) {
                lblStatus.setText("PRODUCT VIEW - MAJOR CASED MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                productDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, false);
            }
            if (prodMJMissedCases == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnTSQueueProd) {
            if (prodQueueTS != 0) {
                lblStatus.setText("PRODUCT VIEW - CASES IN RTS QUEUE");
                tableCases.getItems().clear();
                initTableView(tableCases);
                productViewCasesQueued(tableCases, apnTableView, false);
            }
            if (prodQueueTS == 0) {
                alertUser();
            }
        }

        if (event.getSource() == btnPSQueueProd) {
            if (prodQueuePS != 0) {
                lblStatus.setText("PRODUCT VIEW - CASES IN GPS QUEUE");
                tableCases.getItems().clear();
                initTableView(tableCases);
                productViewCasesQueued(tableCases, apnTableView, true);
            }
            if (prodQueuePS == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnCustomerCritical) {
            if (customerE1 != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Critical";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, true);
            }
            if (customerE1 == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnCustomerE2) {
            if (customerE2 != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Severity";
                String filter = "E2";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, true);
            }
            if (customerE2 == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnCustomerOutFollow) {
            if (customerOutFol != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Outage Follow-Up";
                String filter = "1";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, true);
            }
            if (customerOutFol == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnCustomerEscalated) {
            if (customerEsc != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Escalated By";
                String filter = "NotSet";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, false);
            }
            if (customerEsc == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnCustomerHotIssues) {
            if (customerHot != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Support Hotlist Level";
                String filter = "NotSet";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, false);
            }
            if (customerHot == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnCustomerBC) {
            if (customerBC != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, true);
            }
            if (customerBC == 0) {
                alertUser();
            }
        }
        if (event.getSource() == btnCustomerActWOH) {
            if (customerWoh != 0) {
                tableCustomers.getItems().clear();
                initTableView(tableCustomers);
                customerWOHTable(tableCustomers, true);
            }
            if (customerWoh == 0) {
                alertUser();
                tableCustomers.setVisible(false);
            }
        }
    }

    private void customerWOHTable(TableView<CaseTableView> tableCustomers, boolean bool) {

        int caseCount = 0;

        tableCustomers.setVisible(true);

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
                if (filterColName.equals("Severity")) {
                    caseSevRefCell = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if (!customerText.getText().isEmpty()) {

                ArrayList<String> setCustomerAsItis = new ArrayList<>(Arrays.asList(customerText.getText().split(",\\s*")));
                //ArrayList<String> setCustomerFinal = new ArrayList<>(Arrays.asList(customerText.getText().toUpperCase().split(",\\s*")));
                //ArrayList<String> setCustomerCap = new ArrayList();
                /*int customerNum = setCustomerAsItis.size();

                for (int i = 0; i < customerNum; i++) {

                    Pattern pattern = Pattern.compile("\\b([a-z])([\\w]*)");
                    Matcher matcher = pattern.matcher(setCustomerAsItis.get(i));
                    StringBuffer buffer = new StringBuffer();
                    while (matcher.find()) {
                        matcher.appendReplacement(buffer, matcher.group(1).toUpperCase() + matcher.group(2));
                    }
                    String capitalized = matcher.appendTail(buffer).toString();
                    setCustomerCap.add(capitalized);
                }

                int setcust2num = setCustomerCap.size();
                for (int i = 0; i < setcust2num ; i++) {

                    setCustomerFinal.add(setCustomerCap.get(i));
                }*/
                int customerNumFinal = setCustomerAsItis.size();

                if ((!setCustomerAsItis.isEmpty())) {

                    for (int j = 0; j < customerNumFinal; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(caseAccountRef);
                            String accountName = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();

                            if (bool) {
                                if (accountName.equals(setCustomerAsItis.get(j)) &&
                                        (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability"))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);
                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCustomers.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCustomers.getItems().size() >= caseCount + 1) {
                                        tableCustomers.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if (accountName.equals(setCustomerAsItis.get(j)) &&
                                        (caseStatus.equals("Pending Closure") && caseStatus.equals("Future Availability"))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);
                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCustomers.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCustomers.getItems().size() >= caseCount + 1) {
                                        tableCustomers.getItems().removeAll(observableList);
                                    }
                                }

                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            tableCustomers.setVisible(true);
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCustomers);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCustomers.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCustomers, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCustomers.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCustomers);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });


        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private void customerTable(String columnSelect, String filter, TableView<CaseTableView> tableCustomers, Boolean bool) {

        int caseCount = 0;

        tableCustomers.setVisible(true);

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if (!customerText.getText().isEmpty()) {

                ArrayList<String> setCustomerAsItis = new ArrayList<>(Arrays.asList(customerText.getText().split(",\\s*")));
                //ArrayList<String> setCustomerFinal = new ArrayList<>(Arrays.asList(customerText.getText().toUpperCase().split(",\\s*")));
                //ArrayList<String> setCustomerCap = new ArrayList();
                int customerNum = setCustomerAsItis.size();


                if ((!setCustomerAsItis.isEmpty())) {

                    for (int j = 0; j < customerNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(caseAccountRef);
                            String accountName = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();

                            if (bool) {
                                if (accountName.equals(setCustomerAsItis.get(j)) && cellToCompare.equals(filter) &&
                                        (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability"))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);
                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCustomers.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCustomers.getItems().size() >= caseCount + 1) {
                                        tableCustomers.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if (accountName.equals(setCustomerAsItis.get(j)) && !cellToCompare.equals(filter) &&
                                        (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability"))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);
                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCustomers.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCustomers.getItems().size() >= caseCount + 1) {
                                        tableCustomers.getItems().removeAll(observableList);
                                    }
                                }

                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            tableCustomers.setVisible(true);
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCustomers);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCustomers.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCustomers, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });
            // Selecting and Copy the Case Number to Clipboard
            tableCustomers.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCustomers);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void productViewCasesQueued(TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Support Product")) {
                    caseProductRef = i;
                }
                if (filterColName.equals("Case Owner")) {
                    caseOwnerRefCell = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if (!txProducts.getText().isEmpty()) {

                ArrayList<String> setProd = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));
                int productNum = setProd.size();

                if ((!setProd.isEmpty())) {

                    for (int j = 0; j < productNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(caseProductRef);
                            String productName = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseOwnerRefCell);
                            String owner = cellVal3.getStringCellValue();
                            cellVal4 = filtersheet.getRow(i).getCell(mycaseAgeRefCell);
                            int compAge = Integer.parseInt(cellVal4.getStringCellValue());


                            if (b) {
                                if (productName.equals(setProd.get(j)) && owner.startsWith("PS ")) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);
                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if (productName.equals(setProd.get(j)) && (owner.startsWith("TS ") || owner.startsWith("Tech-Ops"))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnProduct.toFront();
                    lblStatus.setText("PRODUCT VIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void productDueFilterView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, int dueDay, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Support Product")) {
                    caseProductRef = i;
                }
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if (!txProducts.getText().isEmpty()) {

                ArrayList<String> setProd = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));
                int productNum = setProd.size();

                if ((!setProd.isEmpty())) {

                    for (int j = 0; j < productNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(caseProductRef);
                            String productName = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal4 = filtersheet.getRow(i).getCell(mycaseAgeRefCell);
                            int compAge = Integer.parseInt(cellVal4.getStringCellValue());


                            if (b) {
                                if ((productName.equals(setProd.get(j)) && cellToCompare.equals(filter) && compAge < dueDay) &&
                                        (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability"))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if ((productName.equals(setProd.get(j)) && cellToCompare.equals(filter) && compAge > dueDay)
                                        && (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability"))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnProduct.toFront();
                    lblStatus.setText("PRODUCT VIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void inactiveCasesProductTableView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Support Product")) {
                    caseProductRef = i;
                }
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if (!txProducts.getText().isEmpty()) {

                ArrayList<String> setProd = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));

                int productNum = setProd.size();

                if ((!setProd.isEmpty())) {

                    for (int j = 0; j < productNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(caseProductRef);
                            String productName = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();

                            if (productName.equals(setProd.get(j)) && cellToCompare.equals(filter) &&
                                    (caseStatus.equals("Pending Closure") || caseStatus.equals("Future Availability"))) {

                                ArrayList<String> array = new ArrayList<>();
                                ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                while (iterCells.hasNext()) {
                                    HSSFCell cell = (HSSFCell) iterCells.next();
                                    array.add(cell.getStringCellValue());
                                }

                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                LocalDate localDate = null;

                                if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                    localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                }

                                int age = 0;
                                age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), age,
                                        localDate, array.get(7), array.get(8),
                                        array.get(9), array.get(10), array.get(11),
                                        array.get(12), array.get(13), array.get(14),
                                        array.get(15), array.get(16)));

                                tableCases.getItems().addAll(observableList);
                                caseCount++;
                                if (tableCases.getItems().size() >= caseCount + 1) {
                                    tableCases.getItems().removeAll(observableList);
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnProduct.toFront();
                    lblStatus.setText("PRODUCT VIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void twoFilterProductTableView(String columSelect1, String filter1, String columSelect2, String filter2, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Support Product")) {
                    caseProductRef = i;
                }
                if (filterColName.equals(columSelect1)) {
                    caseCellRef = i;
                }
                if (filterColName.equals(columSelect2)) {
                    caseCellRef2 = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if (!txProducts.getText().isEmpty()) {

                ArrayList<String> setProd = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));
                int prodNum = setProd.size();

                if ((!setProd.isEmpty())) {

                    for (int j = 0; j < prodNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(caseProductRef);
                            String productName = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal4 = filtersheet.getRow(i).getCell(caseCellRef2);
                            String responsible = cellVal4.getStringCellValue();


                            if (productName.equals(setProd.get(j)) && cellToCompare.equals(filter1) && responsible.equals(filter2) &&
                                    (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) {

                                ArrayList<String> array = new ArrayList<>();
                                ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                while (iterCells.hasNext()) {
                                    HSSFCell cell = (HSSFCell) iterCells.next();
                                    array.add(cell.getStringCellValue());
                                }

                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                LocalDate localDate = null;

                                if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                    localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                }

                                int age = 0;
                                age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), age,
                                        localDate, array.get(7), array.get(8),
                                        array.get(9), array.get(10), array.get(11),
                                        array.get(12), array.get(13), array.get(14),
                                        array.get(15), array.get(16)));

                                tableCases.getItems().addAll(observableList);
                                caseCount++;
                                if (tableCases.getItems().size() >= caseCount + 1) {
                                    tableCases.getItems().removeAll(observableList);
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnProduct.toFront();
                    lblStatus.setText("PRODUCT VIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void productWIPCaseView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Support Product")) {
                    caseProductRef = i;
                }
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if (!txProducts.getText().isEmpty()) {

                ArrayList<String> setProd = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));

                int productNum = setProd.size();

                if ((!setProd.isEmpty())) {

                    for (int j = 0; j < productNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(caseProductRef);
                            String productName = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();

                            if (productName.equals(setProd.get(j)) && cellToCompare.equals(filter) && (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault")))) {

                                ArrayList<String> array = new ArrayList<>();
                                ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                while (iterCells.hasNext()) {
                                    HSSFCell cell = (HSSFCell) iterCells.next();
                                    array.add(cell.getStringCellValue());
                                }

                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                LocalDate localDate = null;

                                if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                    localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                }

                                int age = 0;
                                age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), age,
                                        localDate, array.get(7), array.get(8),
                                        array.get(9), array.get(10), array.get(11),
                                        array.get(12), array.get(13), array.get(14),
                                        array.get(15), array.get(16)));

                                tableCases.getItems().addAll(observableList);
                                caseCount++;
                                if (tableCases.getItems().size() >= caseCount + 1) {
                                    tableCases.getItems().removeAll(observableList);
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnProduct.toFront();
                    lblStatus.setText("PRODUCT VIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void prodWOHTable(TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Support Product")) {
                    caseProductRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if (!txProducts.getText().isEmpty()) {

                ArrayList<String> setProd = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));

                int prodNumber = setProd.size();

                if ((!setProd.isEmpty())) {

                    for (int j = 0; j < prodNumber; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(caseProductRef);
                            String prodName = cellVal1.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();

                            if (b) {
                                if (prodName.equals(setProd.get(j)) && (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if (prodName.equals(setProd.get(j)) && (caseStatus.equals("Pending Closure") || (caseStatus.equals("Future Availability")))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnProduct.toFront();
                    lblStatus.setText("PRODUCT VIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void productOneFilterView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, Boolean bool) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Support Product")) {
                    caseProductRef = i;
                }
                if (filterColName.equals(columnSelect)) {
                    myCaseCellRef1 = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if ((!txProducts.getText().isEmpty())) {

                ArrayList<String> setProduct = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));

                int setProductCount = setProduct.size();

                if ((!setProduct.isEmpty())) {

                    for (int j = 0; j < setProductCount; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(caseProductRef);
                            String product = cellVal1.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(myCaseCellRef1);
                            String cellValToCompare = cellVal2.getStringCellValue();

                            if (bool) {
                                if ((product.equals(setProduct.get(j)) && cellValToCompare.equals(filter)) && (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if ((product.equals(setProduct.get(j)) && !cellValToCompare.equals(filter)) && (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnProduct.toFront();
                    lblStatus.setText("PRODUCT VIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void downloadCSV() {

        String filename = "cmt_case_data.csv";
        String filename2 = "cmt_user_prod.csv";
        String filename3 = "cmt_survey.csv";

        String newLoc = "https://na8.salesforce.com/00OC0000006r1EX?export=1&enc=UTF-8&xf=csv?filename=" + filename;
        String newLoc2 = "https://na8.salesforce.com/00OC0000006r1xS?export=1&enc=UTF-8&xf=csv?filename=" + filename2;
        String newLoc3 = "https://na8.salesforce.com/00OC0000006r36Q?export=1&enc=UTF-8&xf=csv?filename=" + filename3;
        try {
            FileUtils.copyURLToFile(new URL(newLoc), new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.csv"));
            LocalDate refreshDate = LocalDate.now();
            DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm");
            lblRefreshText.setVisible(true);
            String dataDate = "Data Time Stamp is:" + "\n" + LocalTime.now().format(dtf).toString() + "\n" + "\n" + refreshDate.toString();
            lblRefreshText.setText(dataDate);

            FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_data_Date.txt"));
            writer.write(dataDate);
            writer.close();

        } catch (IOException e) {
            e.printStackTrace();
        }

        try {

            FileUtils.copyURLToFile(new URL(newLoc2), new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_user_prod.csv"));


        } catch (Exception e) {
            e.printStackTrace();
        }

        /*try{

            FileUtils.copyURLToFile(new URL(newLoc3), new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_survey.csv"));

        }catch (Exception e){
            e.printStackTrace();
        }*/

        parseData();
        parseUserData();
        overviewPage();

        time = new Timeline();
        time.setCycleCount(Timeline.INDEFINITE);
        time.getKeyFrames().add(new KeyFrame(Duration.minutes(15), new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                time.stop();
                downloadCSV();
            }
        }));
        time.playFromStart();
    }

    private void connectOkta() {

        if (!btnLogin.getText().equals("Logged!")) {


            browserLoginPane.getChildren().remove(browserLogin);
            WebEngine webEngine = browserLogin.getEngine();
            webEngine.load("https://sonus.okta.com");
            //browserLogin.setPrefSize(1024, 768);
            browserLoginPane.getChildren().add(browserLogin);
            browserLoginPane.toFront();
            apnBrowser.toFront();
            progressBar.setVisible(true);
            progressBar.toFront();
            progressBar.setProgress(0.20);

            webEngine.getLoadWorker().stateProperty().addListener(new ChangeListener<Worker.State>() {
                @Override
                public void changed(ObservableValue ov, Worker.State oldState, Worker.State newState) {

                    if (newState == Worker.State.SUCCEEDED) {
                        if (webEngine.getLocation().equals("https://sonus.okta.com/app/UserHome")) {
                            progressBar.setProgress(0.40);
                            webEngine.load("https://sonus.okta.com/home/salesforce/0oayiqwes0HuzLJ6a1t6/46?fromHome=true");
                            progressBar.setProgress(0.50);
                            apnBrowser.toBack();
                            progressBar.setProgress(0.70);
                        }
                        if (webEngine.getLocation().equals("https://na8.salesforce.com/500/o") || webEngine.getLocation().equals("https://na8.salesforce.com/home/home.jsp")) {
                            progressBar.setProgress(1.00);
                            btnLoadData.setVisible(true);
                            progressBar.setVisible(false);
                            btnLogin.setText("Logged!");
                        }
                    }
                }
            });
        }
    }

    private void projectsPage() {

        //TODO: Project View

    }

    private void writeDefaultSettingsToFile(String userFilter, String queueFilter, String productFilter) {


        ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
        ArrayList<String> setUser2 = new ArrayList();
        int userArraySize = setUser.size();

        for (int i = 0; i < userArraySize; i++) {

            Pattern pattern = Pattern.compile("\\b([a-z])([\\w]*)");
            Matcher matcher = pattern.matcher(setUser.get(i));
            StringBuffer buffer = new StringBuffer();
            while (matcher.find()) {
                matcher.appendReplacement(buffer, matcher.group(1).toUpperCase() + matcher.group(2));
            }
            String capitalized = matcher.appendTail(buffer).toString();
            setUser2.add(capitalized);
        }

        settingsUsers = (ArrayList<String>) setUser2.stream().distinct().collect(Collectors.toList());

        try {

            FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_user_default_settings.txt"));
            int size = settingsUsers.size();
            for (int i = 0; i < size; i++) {
                String str = settingsUsers.get(i);
                writer.write(str);
                if (i < size - 1)
                    writer.write("\n");
            }

            writer.close();

        } catch (Exception e) {
            e.printStackTrace();
        }


        ArrayList<String> setqueue = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
        ArrayList<String> setqueue2 = new ArrayList();
        int queueArraySize = setqueue.size();

        for (int i = 0; i < queueArraySize; i++) {

            Pattern pattern = Pattern.compile("\\b([a-z])([\\w]*)");
            Matcher matcher = pattern.matcher(setqueue.get(i));
            StringBuffer buffer = new StringBuffer();
            while (matcher.find()) {
                matcher.appendReplacement(buffer, matcher.group(1).toUpperCase() + matcher.group(2));
            }
            String capitalized = matcher.appendTail(buffer).toString();
            setqueue2.add(capitalized);
        }

        settingsQueue = (ArrayList<String>) setqueue2.stream().distinct().collect(Collectors.toList());

        try {

            FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_queueu_default_settings.txt"));
            int size = settingsQueue.size();
            for (int i = 0; i < size; i++) {
                String str = settingsQueue.get(i);
                writer.write(str);
                if (i < size - 1)
                    writer.write("\n");
            }
            writer.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

        ArrayList<String> setprod = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));
        ArrayList<String> sfdcProducts = new ArrayList<>();
        int productArraySize = setprod.size();

        for (int i = 0; i < productArraySize; i++) {

            Pattern pattern = Pattern.compile("\\b([a-z])([\\w]*)");
            Matcher matcher = pattern.matcher(setprod.get(i));
            StringBuffer buffer = new StringBuffer();
            while (matcher.find()) {
                matcher.appendReplacement(buffer, matcher.group(1).toUpperCase() + matcher.group(2));
            }
            String capitalized = matcher.appendTail(buffer).toString();
            sfdcProducts.add(capitalized);
        }

        settingsProducts = (ArrayList<String>) sfdcProducts.stream().distinct().collect(Collectors.toList());

        try {

            FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_product_default_settings.txt"));
            int size = settingsProducts.size();
            for (int i = 0; i < size; i++) {
                String str = settingsProducts.get(i);
                writer.write(str);
                if (i < size - 1)
                    writer.write("\n");
            }
            writer.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void readTimeStamp(){

        File timeStampFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_data_Date.txt");

        if (timeStampFile.isFile()){
            Scanner s = null;
            try{

                s = new Scanner(timeStampFile);

            }catch (Exception e){
                e.printStackTrace();
            }

            ArrayList<String> readDate = new ArrayList<>();
            while(s.hasNextLine()){
                readDate.add(s.nextLine());
            }
            s.close();

            System.out.println(readDate);

            lblRefreshText.setVisible(true);
            lblRefreshText.setText(readDate.get(0)+ "\n" + readDate.get(1) + "\n" + readDate.get(2));
        }

    }

    private void readDefaultSettingFiles() {

        // Load Already Saved Settings File if there are any

        File settingUsersFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_user_default_settings.txt");
        File settingQueueFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_queueu_default_settings.txt");
        File settingProductsFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_product_default_settings.txt");

        if (settingUsersFile.isFile()) {

            Scanner s = null;
            try {
                s = new Scanner(settingUsersFile);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            ArrayList<String> readUserList = new ArrayList<String>();
            while (s.hasNextLine()) {
                readUserList.add(s.nextLine());
            }
            s.close();

            txUsers.setText(readUserList.stream().collect(Collectors.joining(", ")));
        }

        if (settingQueueFile.isFile()) {

            Scanner s = null;
            try {
                s = new Scanner(settingQueueFile);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            ArrayList<String> readQueueList = new ArrayList<String>();
            while (s.hasNextLine()) {
                readQueueList.add(s.nextLine());
            }
            s.close();

            txQueues.setText(readQueueList.stream().collect(Collectors.joining(", ")));

        }

        if (settingProductsFile.isFile()) {

            Scanner s = null;
            try {
                s = new Scanner(settingProductsFile);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            ArrayList<String> readProductList = new ArrayList<String>();
            while (s.hasNextLine()) {
                readProductList.add(s.nextLine());
            }
            s.close();

            txProducts.setText(readProductList.stream().collect(Collectors.joining(", ")));

        }
    }

    private void caseUpdateTableView(String caseTableSelect, TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b, boolean bool) {

        int caseCount = 0;

        LocalDate dateToday = LocalDate.now();
        LocalDate caseUpdateDate = null;
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");


        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellValStat;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();
                if (filterColName.equals(caseTableSelect)) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    caseAgeRefCell = i;
                }
            }

            for (int k = 1; k < lastRow + 1; k++) {

                cellVal = filtersheet.getRow(k).getCell(caseNextUpdateDateRef);
                String cellValToCompare = cellVal.getStringCellValue();

                cellValStat = filtersheet.getRow(k).getCell(caseStatRefCell);
                String cellStat = cellValStat.getStringCellValue();

                ArrayList<String> array = new ArrayList<>();
                ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                if (!cellValToCompare.equals("NotSet")) {

                    caseUpdateDate = LocalDate.parse(cellValToCompare, formatter);
                } else {
                    caseUpdateDate = null;
                }

                if ((b) && (bool)) {

                    if ((caseUpdateDate != null)) {

                        if (caseUpdateDate.compareTo(dateToday) == 0) {

                            Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                            while (iterCells.hasNext()) {
                                HSSFCell cell = (HSSFCell) iterCells.next();
                                array.add(cell.getStringCellValue());
                            }

                            int age = 0;
                            age = Integer.parseInt(array.get(caseAgeRefCell));
                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), age,
                                    caseUpdateDate, array.get(7), array.get(8),
                                    array.get(9), array.get(10), array.get(11),
                                    array.get(12), array.get(13), array.get(14),
                                    array.get(15), array.get(16)));

                            tableCases.getItems().addAll(observableList);
                            caseCount++;
                            if (tableCases.getItems().size() >= caseCount + 1) {
                                tableCases.getItems().removeAll(observableList);
                            }
                        }
                    }
                }
                if ((!b) && (bool)) {

                    if ((caseUpdateDate != null)) {

                        if (caseUpdateDate.compareTo(dateToday) < 0) {

                            Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                            while (iterCells.hasNext()) {
                                HSSFCell cell = (HSSFCell) iterCells.next();
                                array.add(cell.getStringCellValue());
                            }

                            int age = 0;
                            age = Integer.parseInt(array.get(caseAgeRefCell));
                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), age,
                                    caseUpdateDate, array.get(7), array.get(8),
                                    array.get(9), array.get(10), array.get(11),
                                    array.get(12), array.get(13), array.get(14),
                                    array.get(15), array.get(16)));

                            tableCases.getItems().addAll(observableList);
                            caseCount++;
                            if (tableCases.getItems().size() >= caseCount + 1) {
                                tableCases.getItems().removeAll(observableList);
                            }
                        }
                    }
                }
                if (!b && !bool) {

                    if ((caseUpdateDate == null)) {

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        int age = 0;
                        age = Integer.parseInt(array.get(caseAgeRefCell));
                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), age,
                                caseUpdateDate, array.get(7), array.get(8),
                                array.get(9), array.get(10), array.get(11),
                                array.get(12), array.get(13), array.get(14),
                                array.get(15), array.get(16)));

                        tableCases.getItems().addAll(observableList);
                        caseCount++;
                        if (tableCases.getItems().size() >= caseCount + 1) {
                            tableCases.getItems().removeAll(observableList);
                        }
                    }
                }

            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();

            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });
            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Select and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }

                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnHome.toFront();
                    lblStatus.setText("GENERAL OVERVIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void mycaseUpdateTableView(String caseTableSelect, TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b, boolean bool) {

        int caseCount = 0;

        LocalDate dateToday = LocalDate.now();
        LocalDate caseUpdateDate = null;
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");


        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellValStat;
            HSSFCell cellValUser;


            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();
                if (filterColName.equals(caseTableSelect)) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    caseAgeRefCell = i;
                }
                if (filterColName.equals("Case Owner")) {
                    mycaseOwnerRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    mycaseUpdateCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty()) || !(txQueues.getText().isEmpty())) {

                ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                ArrayList<String> mergedOwner = new ArrayList<>();

                int userfiltnum = setUser.size();
                int userqueuenum = setQueu.size();

                if (!setUser.isEmpty()) {
                    for (int i = 0; i < userfiltnum; i++) {

                        mergedOwner.add(setUser.get(i));
                    }
                }

                if (!setQueu.isEmpty()) {
                    for (int i = 0; i < userqueuenum; i++) {
                        mergedOwner.add(setQueu.get(i));
                    }
                }

                int mergedUserNum = mergedOwner.size();

                for (int i = 0; i < mergedUserNum; i++) {

                    for (int k = 1; k < lastRow + 1; k++) {

                        cellVal = filtersheet.getRow(k).getCell(caseNextUpdateDateRef);
                        String cellValToCompare = cellVal.getStringCellValue();

                        cellValStat = filtersheet.getRow(k).getCell(caseStatRefCell);
                        String cellStat = cellValStat.getStringCellValue();

                        cellValUser = filtersheet.getRow(k).getCell(mycaseOwnerRefCell);
                        String caseUser = cellValUser.getStringCellValue();

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                        if (!cellValToCompare.equals("NotSet")) {

                            caseUpdateDate = LocalDate.parse(cellValToCompare, formatter);
                        }

                        if ((b) && (bool)) {

                            if ((caseUser.equals(mergedOwner.get(i)) && caseUpdateDate != null)) {

                                if (caseUpdateDate.compareTo(dateToday) == 0) {

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(caseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            caseUpdateDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
                        if ((!b) && (bool)) {

                            if ((caseUser.equals(mergedOwner.get(i)) && caseUpdateDate != null)) {

                                if (caseUpdateDate.compareTo(dateToday) < 0) {

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(caseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            caseUpdateDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
                        if (!b && !bool) {

                            if ((caseUser.equals(mergedOwner.get(i)) && cellValToCompare.equals("NotSet")) && !cellStat.equals("Pending Closure")) {

                                Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                                while (iterCells.hasNext()) {
                                    HSSFCell cell = (HSSFCell) iterCells.next();
                                    array.add(cell.getStringCellValue());
                                }

                                caseUpdateDate = null;

                                int age = 0;
                                age = Integer.parseInt(array.get(caseAgeRefCell));
                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), age,
                                        caseUpdateDate, array.get(7), array.get(8),
                                        array.get(9), array.get(10), array.get(11),
                                        array.get(12), array.get(13), array.get(14),
                                        array.get(15), array.get(16)));

                                tableCases.getItems().addAll(observableList);
                                caseCount++;
                                if (tableCases.getItems().size() >= caseCount + 1) {
                                    tableCases.getItems().removeAll(observableList);
                                }
                            }
                        }

                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();

            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Select and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }

                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnMyCases.toFront();
                    lblStatus.setText("MY CASES");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    private void createMyQueueCaseView(String columnSelect, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();
                if (filterColName.equals(columnSelect)) {
                    myCaseCellRef1 = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
            }

            if (!txQueues.getText().isEmpty()) {

                String queueFilter = txQueues.getText();
                ArrayList<String> setQueue = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                int queuefiltnum = setQueue.size();

                if ((!queueFilter.equals(""))) {

                    for (int j = 0; j < queuefiltnum; j++) {

                        for (int k = 1; k < lastRow + 1; k++) {

                            cellVal = filtersheet.getRow(k).getCell(myCaseCellRef1);
                            String cellValToCompare = cellVal.getStringCellValue();

                            if (cellValToCompare.equals(setQueue.get(j))) {

                                ArrayList<String> array = new ArrayList<>();
                                ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                                while (iterCells.hasNext()) {
                                    HSSFCell cell = (HSSFCell) iterCells.next();
                                    array.add(cell.getStringCellValue());
                                }

                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                LocalDate localDate = null;

                                if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                    localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                }

                                int age;
                                age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), age,
                                        localDate, array.get(7), array.get(8),
                                        array.get(9), array.get(10), array.get(11),
                                        array.get(12), array.get(13), array.get(14),
                                        array.get(15), array.get(16)));

                                tableCases.getItems().addAll(observableList);
                                caseCount++;
                                if (tableCases.getItems().size() >= caseCount + 1) {
                                    tableCases.getItems().removeAll(observableList);
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();

            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnMyCases.toFront();
                    lblStatus.setText("MY CASES");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void oneFilterMyTableView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Case Owner")) {
                    mycaseOwnerRefCell = i;
                }
                if (filterColName.equals(columnSelect)) {
                    myCaseCellRef1 = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty()) || !(txQueues.getText().isEmpty())) {

                ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                ArrayList<String> mergedOwner = new ArrayList<>();

                int userfiltnum = setUser.size();
                int userqueuenum = setQueu.size();

                if (!setUser.isEmpty()) {
                    for (int i = 0; i < userfiltnum; i++) {

                        mergedOwner.add(setUser.get(i));
                    }
                }

                if (!setQueu.isEmpty()) {
                    for (int i = 0; i < userqueuenum; i++) {
                        mergedOwner.add(setQueu.get(i));
                    }
                }

                int mergedUserNum = mergedOwner.size();


                if ((!mergedOwner.isEmpty())) {

                    for (int j = 0; j < mergedUserNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(myCaseCellRef1);
                            String cellValToCompare = cellVal2.getStringCellValue();

                            if (b) {
                                if ((caseUser.equals(mergedOwner.get(j)) && cellValToCompare.equals(filter)) && (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if ((caseUser.equals(mergedOwner.get(j)) && !cellValToCompare.equals(filter)) && (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnMyCases.toFront();
                    lblStatus.setText("MY CASES");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void myWOHTableView(TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Case Owner")) {
                    mycaseOwnerRefCell = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty()) || !(txQueues.getText().isEmpty())) {

                ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                ArrayList<String> mergedOwner = new ArrayList<>();

                int userfiltnum = setUser.size();
                int userqueuenum = setQueu.size();

                if (!setUser.isEmpty()) {
                    for (int i = 0; i < userfiltnum; i++) {

                        mergedOwner.add(setUser.get(i));
                    }
                }

                if (!setQueu.isEmpty()) {
                    for (int i = 0; i < userqueuenum; i++) {
                        mergedOwner.add(setQueu.get(i));
                    }
                }

                int mergedUserNum = mergedOwner.size();


                if ((!mergedOwner.isEmpty())) {

                    for (int j = 0; j < mergedUserNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();

                            if (b) {
                                if (caseUser.equals(mergedOwner.get(j)) && (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if (caseUser.equals(mergedOwner.get(j)) && (caseStatus.equals("Pending Closure") || (caseStatus.equals("Future Availability")))) {

                                    ArrayList<String> array = new ArrayList<>();
                                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    LocalDate localDate = null;

                                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                    }

                                    int age;
                                    age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), age,
                                            localDate, array.get(7), array.get(8),
                                            array.get(9), array.get(10), array.get(11),
                                            array.get(12), array.get(13), array.get(14),
                                            array.get(15), array.get(16)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnMyCases.toFront();
                    lblStatus.setText("MY CASES");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void overviewMyWIPCaseTableView(String columFilter, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Case Owner")) {
                    mycaseOwnerRefCell = i;
                }
                if (filterColName.equals(columFilter)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty()) || !(txQueues.getText().isEmpty())) {

                ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                ArrayList<String> mergedOwner = new ArrayList<>();

                int userfiltnum = setUser.size();
                int userqueuenum = setQueu.size();

                if (!setUser.isEmpty()) {
                    for (int i = 0; i < userfiltnum; i++) {

                        mergedOwner.add(setUser.get(i));
                    }
                }

                if (!setQueu.isEmpty()) {
                    for (int i = 0; i < userqueuenum; i++) {
                        mergedOwner.add(setQueu.get(i));
                    }
                }

                int mergedUserNum = mergedOwner.size();


                if ((!mergedOwner.isEmpty())) {

                    for (int j = 0; j < mergedUserNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();

                            if (caseUser.equals(mergedOwner.get(j)) && cellToCompare.equals(filter) && (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault")))) {

                                ArrayList<String> array = new ArrayList<>();
                                ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                while (iterCells.hasNext()) {
                                    HSSFCell cell = (HSSFCell) iterCells.next();
                                    array.add(cell.getStringCellValue());
                                }

                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                LocalDate localDate = null;

                                if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                    localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                }

                                int age = 0;
                                age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), age,
                                        localDate, array.get(7), array.get(8),
                                        array.get(9), array.get(10), array.get(11),
                                        array.get(12), array.get(13), array.get(14),
                                        array.get(15), array.get(16)));

                                tableCases.getItems().addAll(observableList);
                                caseCount++;
                                if (tableCases.getItems().size() >= caseCount + 1) {
                                    tableCases.getItems().removeAll(observableList);
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });
            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnMyCases.toFront();
                    lblStatus.setText("MY CASES");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void twoFilterMyTableView(String columSelect1, String filter1, String columSelect2, String filter2, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Case Owner")) {
                    mycaseOwnerRefCell = i;
                }
                if (filterColName.equals(columSelect1)) {
                    caseCellRef = i;
                }
                if (filterColName.equals(columSelect2)) {
                    caseCellRef2 = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty()) || !(txQueues.getText().isEmpty())) {

                ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                ArrayList<String> mergedOwner = new ArrayList<>();

                int userfiltnum = setUser.size();
                int userqueuenum = setQueu.size();

                if (!setUser.isEmpty()) {
                    for (int i = 0; i < userfiltnum; i++) {

                        mergedOwner.add(setUser.get(i));
                    }
                }

                if (!setQueu.isEmpty()) {
                    for (int i = 0; i < userqueuenum; i++) {
                        mergedOwner.add(setQueu.get(i));
                    }
                }

                int mergedUserNum = mergedOwner.size();


                if ((!mergedOwner.isEmpty())) {

                    for (int j = 0; j < mergedUserNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal4 = filtersheet.getRow(i).getCell(caseCellRef2);
                            String responsible = cellVal4.getStringCellValue();


                            if (caseUser.equals(mergedOwner.get(j)) && cellToCompare.equals(filter1) && responsible.equals(filter2) &&
                                    (!caseStatus.equals("Pending Closure") || (!caseStatus.equals("Future Availability")))) {

                                ArrayList<String> array = new ArrayList<>();
                                ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                while (iterCells.hasNext()) {
                                    HSSFCell cell = (HSSFCell) iterCells.next();
                                    array.add(cell.getStringCellValue());
                                }

                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                LocalDate localDate = null;

                                if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                    localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                }

                                int age = 0;
                                age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), age,
                                        localDate, array.get(7), array.get(8),
                                        array.get(9), array.get(10), array.get(11),
                                        array.get(12), array.get(13), array.get(14),
                                        array.get(15), array.get(16)));

                                tableCases.getItems().addAll(observableList);
                                caseCount++;
                                if (tableCases.getItems().size() >= caseCount + 1) {
                                    tableCases.getItems().removeAll(observableList);
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });
            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnMyCases.toFront();
                    lblStatus.setText("MY CASES");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void inactiveCasesMyTableView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Case Owner")) {
                    mycaseOwnerRefCell = i;
                }
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty()) || !(txQueues.getText().isEmpty())) {

                ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                ArrayList<String> mergedOwner = new ArrayList<>();

                int userfiltnum = setUser.size();
                int userqueuenum = setQueu.size();

                if (!setUser.isEmpty()) {
                    for (int i = 0; i < userfiltnum; i++) {

                        mergedOwner.add(setUser.get(i));
                    }
                }

                if (!setQueu.isEmpty()) {
                    for (int i = 0; i < userqueuenum; i++) {
                        mergedOwner.add(setQueu.get(i));
                    }
                }

                int mergedUserNum = mergedOwner.size();


                if ((!mergedOwner.isEmpty())) {

                    for (int j = 0; j < mergedUserNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();

                            if (caseUser.equals(mergedOwner.get(j)) && cellToCompare.equals(filter) &&
                                    (caseStatus.equals("Pending Closure") || caseStatus.equals("Future Availability"))) {

                                ArrayList<String> array = new ArrayList<>();
                                ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                while (iterCells.hasNext()) {
                                    HSSFCell cell = (HSSFCell) iterCells.next();
                                    array.add(cell.getStringCellValue());
                                }

                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                LocalDate localDate = null;

                                if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                    localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                }

                                int age = 0;
                                age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), age,
                                        localDate, array.get(7), array.get(8),
                                        array.get(9), array.get(10), array.get(11),
                                        array.get(12), array.get(13), array.get(14),
                                        array.get(15), array.get(16)));

                                tableCases.getItems().addAll(observableList);
                                caseCount++;
                                if (tableCases.getItems().size() >= caseCount + 1) {
                                    tableCases.getItems().removeAll(observableList);
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnMyCases.toFront();
                    lblStatus.setText("MY CASES");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    private void myDueFilterView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, int dueDay, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Case Owner")) {
                    mycaseOwnerRefCell = i;
                }
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    mycaseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty()) || !(txQueues.getText().isEmpty())) {

                ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                ArrayList<String> mergedOwner = new ArrayList<>();

                int userfiltnum = setUser.size();
                int userqueuenum = setQueu.size();

                if (!setUser.isEmpty()) {
                    for (int i = 0; i < userfiltnum; i++) {

                        mergedOwner.add(setUser.get(i));
                    }
                }

                if (!setQueu.isEmpty()) {
                    for (int i = 0; i < userqueuenum; i++) {
                        mergedOwner.add(setQueu.get(i));
                    }
                }

                int mergedUserNum = mergedOwner.size();

                if ((!mergedOwner.isEmpty())) {

                    for (int j = 0; j < mergedUserNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal4 = filtersheet.getRow(i).getCell(mycaseAgeRefCell);
                            int compAge = Integer.parseInt(cellVal4.getStringCellValue());

                            if ((cellToCompare.equals(filter))) {

                                if (b) {

                                    if ((caseUser.equals(mergedOwner.get(j)) && (compAge <= dueDay)) && ((caseStatus.equals("Open / Assign") ||
                                            (caseStatus.equals("Isolate Fault"))))) {

                                        ArrayList<String> array = new ArrayList<>();
                                        ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                        while (iterCells.hasNext()) {
                                            HSSFCell cell = (HSSFCell) iterCells.next();
                                            array.add(cell.getStringCellValue());
                                        }

                                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                        LocalDate localDate = null;

                                        if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                            localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                        }

                                        int age = 0;
                                        age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), age,
                                                localDate, array.get(7), array.get(8),
                                                array.get(9), array.get(10), array.get(11),
                                                array.get(12), array.get(13), array.get(14),
                                                array.get(15), array.get(16)));

                                        tableCases.getItems().addAll(observableList);
                                        caseCount++;
                                        if (tableCases.getItems().size() >= caseCount + 1) {
                                            tableCases.getItems().removeAll(observableList);
                                        }
                                    }
                                } else {
                                    if ((caseUser.equals(mergedOwner.get(j)) && (compAge > dueDay)) && ((caseStatus.equals("Open / Assign") ||
                                            (caseStatus.equals("Isolate Fault"))))) {

                                        ArrayList<String> array = new ArrayList<>();
                                        ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(i).cellIterator();
                                        while (iterCells.hasNext()) {
                                            HSSFCell cell = (HSSFCell) iterCells.next();
                                            array.add(cell.getStringCellValue());
                                        }

                                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                        LocalDate localDate = null;

                                        if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                                            localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                                        }

                                        int age;
                                        age = Integer.parseInt(array.get(mycaseAgeRefCell));
                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), age,
                                                localDate, array.get(7), array.get(8),
                                                array.get(9), array.get(10), array.get(11),
                                                array.get(12), array.get(13), array.get(14),
                                                array.get(15), array.get(16)));

                                        tableCases.getItems().addAll(observableList);
                                        caseCount++;
                                        if (tableCases.getItems().size() >= caseCount + 1) {
                                            tableCases.getItems().removeAll(observableList);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });
            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnMyCases.toFront();
                    lblStatus.setText("MY CASES");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void oneFilterTableView(String columnSelect, String filter1, TableView tableCases, AnchorPane apnTableView, Boolean bool) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    caseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();
                cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal2.getStringCellValue();

                if (!bool) {
                    if (!cellValToCompare.equals(filter1) && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                        LocalDate localDate = null;

                        if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                            localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                        }

                        int age = 0;
                        age = Integer.parseInt(array.get(caseAgeRefCell));
                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), age,
                                localDate, array.get(7), array.get(8),
                                array.get(9), array.get(10), array.get(11),
                                array.get(12), array.get(13), array.get(14),
                                array.get(15), array.get(16)));

                        tableCases.getItems().addAll(observableList);
                        caseCount++;
                        if (tableCases.getItems().size() >= caseCount + 1) {
                            tableCases.getItems().removeAll(observableList);
                        }
                    }
                } else {
                    if (cellValToCompare.equals(filter1) && (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability"))) {
                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                        LocalDate localDate = null;

                        if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                            localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                        }

                        int age = 0;
                        age = Integer.parseInt(array.get(caseAgeRefCell));
                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), age,
                                localDate, array.get(7), array.get(8),
                                array.get(9), array.get(10), array.get(11),
                                array.get(12), array.get(13), array.get(14),
                                array.get(15), array.get(16)));

                        tableCases.getItems().addAll(observableList);
                        caseCount++;
                        if (tableCases.getItems().size() >= caseCount + 1) {
                            tableCases.getItems().removeAll(observableList);
                        }
                    }
                }

                btnToExcel.setVisible(true);
                apnTableView.toFront();

                btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        exportExcelAction(tableCases);
                    }
                });

                menu = new ContextMenu();
                String caseno = "";
                menu.getItems().add(openCaseSFDC);
                tableCases.setContextMenu(menu);

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            e.printStackTrace();
                        }

                    }
                });

                // Selecting and Copy the Case Number to Clipboard
                tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        try {
                            copyCaseNumberToClipboard(tableCases);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnHome.toFront();
                        lblStatus.setText("GENERAL OVERVIEW");
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        ;
    }

    private void twoFilterTableView(String columnSelect1, String columnSelect2, String filter1, String filter2, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();
                if (filterColName.equals(columnSelect1)) {
                    caseCellRef = i;
                }
                if (filterColName.equals(columnSelect2)) {
                    caseCellRef2 = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    caseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();
                cellVal2 = filtersheet.getRow(k).getCell(caseCellRef2);
                String cellValToCompare2 = cellVal2.getStringCellValue();

                cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal2.getStringCellValue();

                if ((cellValToCompare.equals(filter1) && cellValToCompare2.equals(filter2)) && ((!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")))) {
                    ArrayList<String> array = new ArrayList<>();
                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                    while (iterCells.hasNext()) {
                        HSSFCell cell = (HSSFCell) iterCells.next();
                        array.add(cell.getStringCellValue());
                    }

                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                    LocalDate localDate = null;

                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                    }

                    int age = 0;
                    age = Integer.parseInt(array.get(caseAgeRefCell));
                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), age,
                            localDate, array.get(7), array.get(8),
                            array.get(9), array.get(10), array.get(11),
                            array.get(12), array.get(13), array.get(14),
                            array.get(15), array.get(16)));

                    tableCases.getItems().addAll(observableList);
                    caseCount++;
                    if (tableCases.getItems().size() >= caseCount + 1) {
                        tableCases.getItems().removeAll(observableList);
                    }
                }

                btnToExcel.setVisible(true);
                apnTableView.toFront();

                btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        exportExcelAction(tableCases);
                    }
                });

                menu = new ContextMenu();
                String caseno = "";
                menu.getItems().add(openCaseSFDC);
                tableCases.setContextMenu(menu);

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            e.printStackTrace();
                        }

                    }
                });

                // Selecting and Copy the Case Number to Clipboard
                tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        try {
                            copyCaseNumberToClipboard(tableCases);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnHome.toFront();
                        lblStatus.setText("GENERAL OVERVIEW");
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        ;
    }

    private void overviewWIPCaseTableView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    caseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();

                cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal2.getStringCellValue();

                if (cellValToCompare.equals(filter) && (caseStatus.equals("Open / Assign") || caseStatus.equals("Isolate Fault")) && (!caseStatus.equals("Pending Closure") || !caseStatus.equals("Future Availability"))) {
                    ArrayList<String> array = new ArrayList<>();
                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                    while (iterCells.hasNext()) {
                        HSSFCell cell = (HSSFCell) iterCells.next();
                        array.add(cell.getStringCellValue());
                    }

                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                    LocalDate localDate = null;

                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                    }

                    int age = 0;
                    age = Integer.parseInt(array.get(caseAgeRefCell));
                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), age,
                            localDate, array.get(7), array.get(8),
                            array.get(9), array.get(10), array.get(11),
                            array.get(12), array.get(13), array.get(14),
                            array.get(15), array.get(16)));

                    tableCases.getItems().addAll(observableList);
                    caseCount++;
                    if (tableCases.getItems().size() >= caseCount + 1) {
                        tableCases.getItems().removeAll(observableList);
                    }
                }

                btnToExcel.setVisible(true);
                apnTableView.toFront();

                btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        exportExcelAction(tableCases);
                    }
                });

                menu = new ContextMenu();
                String caseno = "";
                menu.getItems().add(openCaseSFDC);
                tableCases.setContextMenu(menu);

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            e.printStackTrace();
                        }

                    }
                });

                // Selecting and Copy the Case Number to Clipboard
                tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        try {
                            copyCaseNumberToClipboard(tableCases);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnHome.toFront();
                        lblStatus.setText("GENERAL OVERVIEW");
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        ;
    }

    private void overviewEngineeringTableView(String columnSelect, String filter1, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    caseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();

                cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal2.getStringCellValue();

                if (cellValToCompare.equals(filter1) && caseStatus.equals("Develop Solution") && (!caseStatus.equals("Pending Closure") || !caseStatus.equals("Future Availability"))) {
                    ArrayList<String> array = new ArrayList<>();
                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                    while (iterCells.hasNext()) {
                        HSSFCell cell = (HSSFCell) iterCells.next();
                        array.add(cell.getStringCellValue());
                    }

                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                    LocalDate localDate = null;

                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                    }

                    int age = 0;
                    age = Integer.parseInt(array.get(caseAgeRefCell));
                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), age,
                            localDate, array.get(7), array.get(8),
                            array.get(9), array.get(10), array.get(11),
                            array.get(12), array.get(13), array.get(14),
                            array.get(15), array.get(16)));

                    tableCases.getItems().addAll(observableList);
                    caseCount++;
                    if (tableCases.getItems().size() >= caseCount + 1) {
                        tableCases.getItems().removeAll(observableList);
                    }
                }

                btnToExcel.setVisible(true);
                apnTableView.toFront();

                btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        exportExcelAction(tableCases);
                    }
                });

                menu = new ContextMenu();
                String caseno = "";
                menu.getItems().add(openCaseSFDC);
                tableCases.setContextMenu(menu);

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            e.printStackTrace();
                        }

                    }
                });

                // Selecting and Copy the Case Number to Clipboard
                tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        try {
                            copyCaseNumberToClipboard(tableCases);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnHome.toFront();
                        lblStatus.setText("GENERAL OVERVIEW");
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        ;
    }

    private void overViewInactiveTable(String columnSelect1, String filter1, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();
                if (filterColName.equals(columnSelect1)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    caseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();
                cellVal2 = filtersheet.getRow(k).getCell(caseCellRef2);
                String cellValToCompare2 = cellVal2.getStringCellValue();

                cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal2.getStringCellValue();

                if (cellValToCompare.equals(filter1) && (caseStatus.equals("Pending Closure") || caseStatus.equals("Future Availability"))) {
                    ArrayList<String> array = new ArrayList<>();
                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                    while (iterCells.hasNext()) {
                        HSSFCell cell = (HSSFCell) iterCells.next();
                        array.add(cell.getStringCellValue());
                    }

                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                    LocalDate localDate = null;

                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                    }

                    int age = 0;
                    age = Integer.parseInt(array.get(caseAgeRefCell));
                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), age,
                            localDate, array.get(7), array.get(8),
                            array.get(9), array.get(10), array.get(11),
                            array.get(12), array.get(13), array.get(14),
                            array.get(15), array.get(16)));

                    tableCases.getItems().addAll(observableList);
                    caseCount++;
                    if (tableCases.getItems().size() >= caseCount + 1) {
                        tableCases.getItems().removeAll(observableList);
                    }
                }

                btnToExcel.setVisible(true);
                apnTableView.toFront();

                btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        exportExcelAction(tableCases);
                    }
                });

                menu = new ContextMenu();
                String caseno = "";
                menu.getItems().add(openCaseSFDC);
                tableCases.setContextMenu(menu);

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            e.printStackTrace();
                        }

                    }
                });

                // Selecting and Copy the Case Number to Clipboard
                tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        try {
                            copyCaseNumberToClipboard(tableCases);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnHome.toFront();
                        lblStatus.setText("GENERAL OVERVIEW");
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        ;

    }

    private void overviewWOHView(Boolean bool) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Age (Days)")) {
                    caseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();
                cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal2.getStringCellValue();

                if (!bool) {

                    if (caseStatus.equals("Pending Closure") || caseStatus.equals("Future Availability")) {
                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                        LocalDate localDate = null;

                        if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                            localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                        }

                        int age = 0;
                        age = Integer.parseInt(array.get(caseAgeRefCell));
                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), age,
                                localDate, array.get(7), array.get(8),
                                array.get(9), array.get(10), array.get(11),
                                array.get(12), array.get(13), array.get(14),
                                array.get(15), array.get(16)));

                        tableCases.getItems().addAll(observableList);
                        caseCount++;
                        if (tableCases.getItems().size() >= caseCount + 1) {
                            tableCases.getItems().removeAll(observableList);
                        }
                    }

                } else {

                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                        LocalDate localDate = null;

                        if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                            localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                        }

                        int age = 0;
                        age = Integer.parseInt(array.get(caseAgeRefCell));
                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), age,
                                localDate, array.get(7), array.get(8),
                                array.get(9), array.get(10), array.get(11),
                                array.get(12), array.get(13), array.get(14),
                                array.get(15), array.get(16)));

                        tableCases.getItems().addAll(observableList);
                        caseCount++;
                        if (tableCases.getItems().size() >= caseCount + 1) {
                            tableCases.getItems().removeAll(observableList);
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();

            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnHome.toFront();
                    lblStatus.setText("GENERAL OVERVIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            e.printStackTrace();
        }
        ;
    }

    private void overviewQueueView(String columnSelect, String filter, TableView tableView, AnchorPane anchorpane, String overText) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    caseCellRef2 = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }

            }
            for (int k = 1; k < lastRow + 1; k++) {

                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();

                if (cellValToCompare.equals(filter) || cellValToCompare.startsWith(filter)) {

                    ArrayList<String> array = new ArrayList<>();
                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                    while (iterCells.hasNext()) {
                        HSSFCell cell = (HSSFCell) iterCells.next();
                        array.add(cell.getStringCellValue());
                    }

                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                    LocalDate localDate = null;

                    if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                        localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);
                    }

                    int age;
                    age = Integer.parseInt(array.get(caseCellRef2));

                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), age,
                            localDate, array.get(7), array.get(8),
                            array.get(9), array.get(10), array.get(11),
                            array.get(12), array.get(13), array.get(14),
                            array.get(15), array.get(16)));

                    tableView.getItems().addAll(observableList);
                    caseCount++;
                    if (tableView.getItems().size() >= caseCount + 1) {
                        tableView.getItems().removeAll(observableList);
                    }
                }
            }
            btnToExcel.setVisible(true);
            anchorpane.toFront();

            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableView);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableView.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableView);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnHome.toFront();
                    lblStatus.setText("GENERAL OVERVIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            e.printStackTrace();
        }
        ;
    }

    private void overviewDueFilterView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, int ageDue, Boolean due) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();
                if (filterColName.equals(columnSelect)) {
                    caseCellRef = i;
                }
                if (filterColName.equals("Age (Days)")) {
                    caseAgeRefCell = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Status")) {
                    caseStatRefCell = i;
                }
            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();
                cellVal2 = filtersheet.getRow(k).getCell(caseAgeRefCell);
                String caseAge = cellVal2.getStringCellValue();
                cellVal3 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal3.getStringCellValue();
                int ageCase = Integer.parseInt(caseAge);

                if (due) {

                    if ((cellValToCompare.equals(filter) && ageCase <= ageDue) && ((!caseStatus.equals("Develop Solution") &&
                            (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))))) {

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                        LocalDate localDate = null;

                        if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                            localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                        }
                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), ageCase,
                                localDate, array.get(7), array.get(8),
                                array.get(9), array.get(10), array.get(11),
                                array.get(12), array.get(13), array.get(14),
                                array.get(15), array.get(16)));

                        tableCases.getItems().addAll(observableList);
                        caseCount++;
                        if (tableCases.getItems().size() >= caseCount + 1) {
                            tableCases.getItems().removeAll(observableList);
                        }
                    }
                } else {
                    if ((cellValToCompare.equals(filter) && ageCase > ageDue) && ((!caseStatus.equals("Develop Solution") &&
                            (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))))) {

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                        LocalDate localDate = null;

                        if (!array.get(caseNextUpdateDateRef).equals("NotSet")) {

                            localDate = LocalDate.parse(array.get(caseNextUpdateDateRef), formatter);

                        }
                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), ageCase,
                                localDate, array.get(7), array.get(8),
                                array.get(9), array.get(10), array.get(11),
                                array.get(12), array.get(13), array.get(14),
                                array.get(15), array.get(16)));

                        tableCases.getItems().addAll(observableList);
                        caseCount++;
                        if (tableCases.getItems().size() >= caseCount + 1) {
                            tableCases.getItems().removeAll(observableList);
                        }
                    }

                }
            }

            btnToExcel.setVisible(true);
            apnTableView.toFront();

            btnToExcel.setVisible(true);
            apnTableView.toFront();

            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelAction(tableCases);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            tableCases.setContextMenu(menu);

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://na8.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        e.printStackTrace();
                    }

                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnHome.toFront();
                    lblStatus.setText("GENERAL OVERVIEW");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void extractToExcel(TableView tableView, String textData, File file) throws IOException {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet spreadsheet = workbook.createSheet(textData);
        HSSFRow row = spreadsheet.createRow(0);
        TableColumn tableColumn = new TableColumn();

        ArrayList headerArray = new ArrayList();
        headerArray.add("Case Number");
        headerArray.add("Severity");
        headerArray.add("Status");
        headerArray.add("Owner");
        headerArray.add("Responsible");
        headerArray.add("Age");
        headerArray.add("Next Update Date");
        headerArray.add("Escalated By");
        headerArray.add("Hot List");
        headerArray.add("Outage Folllow-up");
        headerArray.add("Support Type");
        headerArray.add("Product");
        headerArray.add("Subject");
        headerArray.add("Account");
        headerArray.add("Region");
        headerArray.add("Security");
        headerArray.add("Date Opened");

        for (int k = 0; k < tableView.getColumns().size(); k++) {
            row.createCell(k).setCellValue(headerArray.get(k).toString());
        }

        for (int i = 0; i < tableView.getItems().size(); i++) {

            row = spreadsheet.createRow(i + 1);

            for (int j = 0; j < tableView.getColumns().size(); j++) {

                tableColumn = (TableColumn) tableView.getColumns().get(j);
                if (tableColumn.getCellObservableValue(i).getValue() != null) {
                    row.createCell(j).setCellValue(tableColumn.getCellObservableValue(i).getValue().toString());
                } else {
                    row.createCell(j).setCellValue("");
                }
            }
        }

        FileOutputStream fileOut = new FileOutputStream(file);
        workbook.write(fileOut);
        fileOut.close();
    }

    /* Creating XLS File from CSV File downloaded*/
    private void parseData() {

        try {
            File csvfile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.csv");

            HSSFWorkbook workBook = new HSSFWorkbook();

            String xlsFileAddress = System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls";
            HSSFSheet sheet = workBook.createSheet("Data");

            BufferedReader br = new BufferedReader(new FileReader(csvfile));
            String line;

            int RowNum = 0;

            while ((line = br.readLine()) != null) {
                line = line.replaceAll("^\"|\"$", "");
                line = line.replaceAll(".0000000000", "");

                String[] fields = parseCsvLine(line);

                HSSFRow currentRow = sheet.createRow(RowNum);
                for (int i = 0; i < fields.length; i++) {
                    currentRow.createCell(i).setCellValue(fields[i]);
                    if (currentRow.getCell(i).getStringCellValue().isEmpty() || currentRow.getCell(i).getStringCellValue().equals("FALSE")) {
                        currentRow.getCell(i).setCellValue("NotSet");
                    }
                }
                RowNum++;
            }

            int lastRow = sheet.getLastRowNum();

            for (int i = 0; i < 7; i++) {
                sheet.removeRow(sheet.getRow(lastRow - i));
            }
            FileOutputStream fileOutputStream = new FileOutputStream(xlsFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void parseUserData() {
        try {
            File csvfile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_user_prod.csv");

            HSSFWorkbook workBook = new HSSFWorkbook();

            String xlsFileAddress = System.getProperty("user.home") + "\\Documents\\CMT\\cmt_user_prod.xls";
            HSSFSheet sheet = workBook.createSheet("UserProd");

            BufferedReader br = new BufferedReader(new FileReader(csvfile));
            String line;

            int RowNum = 0;

            while ((line = br.readLine()) != null) {
                //line = line.replace("Ã©n", "e");
                line = line.replaceAll("^\"|\"$", "");

                String[] fields = parseCsvLine(line);

                HSSFRow currentRow = sheet.createRow(RowNum);
                for (int i = 0; i < fields.length; i++) {
                    currentRow.createCell(i).setCellValue(fields[i]);
                }
                RowNum++;
            }

            int lastRow = sheet.getLastRowNum();

            for (int i = 0; i < 7; i++) {
                sheet.removeRow(sheet.getRow(lastRow - i));
            }
            FileOutputStream fileOutputStream = new FileOutputStream(xlsFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();


        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void parseSurveyData(){

        try {
            File csvfile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_survey.csv");

            HSSFWorkbook workBook = new HSSFWorkbook();

            String xlsFileAddress = System.getProperty("user.home") + "\\Documents\\CMT\\cmt_survey.xls";
            HSSFSheet sheet = workBook.createSheet("Survey");

            BufferedReader br = new BufferedReader(new FileReader(csvfile));
            String line;

            int RowNum = 0;

            while ((line = br.readLine()) != null) {
                line = line.replaceAll("^\"|\"$", "");
                line = line.replace("-&amp;gt;", "");
                line = line.replace("- ", "");
                line = line.replace("&gt;", "");
                line = line.replace("Ribbon&apos;s", "Ribbon's");
                line = line.replace("...", "");
                line = line.replace("Very SATISFIED 10", "10");
                line = line.replace("Strongly AGREE 10", "10");
                line = line.replace("Not Answered", "NA");
                line = line.replace("? Time i", "?");
                line = line.replace("? Profes", "?");
                line = line.replace("? Expertise", "?");
                line = line.replace("? Freque", "?");
                line = line.replace("? On clo", "?");
                line = line.replace("Ribbon Technical Support Customer Survey v2", "Technical Support");
                line = line.replace("Ribbon Technical Support Customer Survey", "Technical Support");
                line = line.replace("Ribbon KBS Support Customer Survey v2", "KBS Support");
                line = line.replace("Ribbon KBS Support Customer Survey", "KBS Support");
                line = line.replace("Ribbon Emergency Recovery Customer Survey v2", "Emergency Recovery Support");
                line = line.replace("Ribbon Emergency Recovery Customer Survey", "Emergency Recovery Support");





                String[] fields = parseCsvLine(line);

                HSSFRow currentRow = sheet.createRow(RowNum);
                for (int i = 0; i < fields.length; i++) {

                    currentRow.createCell(i).setCellValue(fields[i]);
                    if (currentRow.getCell(i).getStringCellValue().isEmpty()) {
                        currentRow.getCell(i).setCellValue("NotSet");
                    }
                }
                RowNum++;

            }
            int lastRow = sheet.getLastRowNum();
            for (int i = 0; i < 7; i++) {
                sheet.removeRow(sheet.getRow(lastRow - i));
            }
            FileOutputStream fileOutputStream = new FileOutputStream(xlsFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    // Splitting the CSV File for reading
    private String[] parseCsvLine(String line) throws IOException {

        Pattern p = Pattern.compile("\",\"");
        //Pattern p = Pattern.compile(",(?=([^\"]*\"[^\"]*\")*(?![^\"]*\"))");

        // Split input with the pattern
        String[] fields = p.split(line);

        return fields;
    }

    private void readSurveyData(){

        ArrayList<String> techCaseNumStr = new ArrayList<>();
        ArrayList<String> kbsCaseNumStr = new ArrayList<>();
        ArrayList<String> erCaseNumStr = new ArrayList<>();


        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_survey.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("RelatedCase: Case Number")){
                    caseNumCellRef = i;
                }
                if (filterColName.equals("RelatedCase: Severity")){
                    caseSevRefCell = i;
                }
                if (filterColName.equals("Question Summary Name")){
                    caseMainQuestionRefCell = i;
                }
                if (filterColName.equals("Sub Question")){
                    caseSubQuestionRefCell = i;
                }
                if (filterColName.equals("Answer(s)")) {
                    caseAnswerRefCell = i;
                }
                if (filterColName.equals("Complete Survey Name")){
                    caseSurveyTypeRef = i;
                }
            }

            HSSFCell caseNum;
            HSSFCell caseNum2;
            HSSFCell caseMain;
            HSSFCell caseSub;
            HSSFCell caseAns;
            HSSFCell surType;

            for (int i = 1; i < lastRow + 1; i++) {

                caseNum = filtersheet.getRow(i).getCell(caseNumCellRef);
                String caseNumber = caseNum.getStringCellValue();

                surType =  filtersheet.getRow(i).getCell(caseSurveyTypeRef);
                String surveyType = surType.getStringCellValue();

                if (surveyType.equals("Technical Support")){

                    techCaseNumStr.add(caseNumber);
                }

                if (surveyType.equals("KBS Support")){

                    kbsCaseNumStr.add(caseNumber);
                }

                if (surveyType.equals("Emergency Recovery Support")){

                    erCaseNumStr.add(caseNumber);
                }

            }

            techCaseNumStr = (ArrayList) techCaseNumStr.stream().distinct().collect(Collectors.toList());
            int techCaseNumStrSize = techCaseNumStr.size();

            kbsCaseNumStr = (ArrayList) kbsCaseNumStr.stream().distinct().collect(Collectors.toList());
            int kbsCaseNumStrSize = kbsCaseNumStr.size();

            erCaseNumStr = (ArrayList) erCaseNumStr.stream().distinct().collect(Collectors.toList());
            int erCaseNumStrSize = techCaseNumStr.size();



           for (int i = 1; i < techCaseNumStrSize + 1; i++) {

               for (int j = 0; j < lastRow ; j++) {
                   
               }

                caseMain = filtersheet.getRow(i).getCell(caseMainQuestionRefCell);
                String casMain = caseMain.getStringCellValue();

            }

        }catch (Exception e){
            e.printStackTrace();
        }

    }


    private void overviewPage() {

        HSSFCell filtHotList;
        HSSFCell outfollow;
        HSSFCell escCases;
        HSSFCell caseSev;
        HSSFCell caseStat;
        HSSFCell ageCase;
        HSSFCell curResp;
        HSSFCell caseOwner;
        HSSFCell caseUpdate;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {
            HSSFSheet filtersheet = workbook.getSheetAt(0);

            hotlist = 0;
            escCase = 0;
            bcCases = 0;
            inactiveCases = 0;
            wohCases = 0;
            bcDue = 0;
            misBCdue = 0;
            custActBC = 0;
            custRpdBC = 0;
            BCds = 0;
            BCpc = 0;
            BCwip = 0;
            dueMJday = 0;
            misMJdue = 0;
            misMNdue = 0;
            custActMJ = 0;
            custRpdMJ = 0;
            MJds = 0;
            MJpc = 0;
            MJwip = 0;
            e1Cases = 0;
            e2Cases = 0;
            outFollow = 0;
            queuePS = 0;
            queueTS = 0;
            updateMissed = 0;
            updateNull = 0;
            updateToday = 0;

            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                switch (filterColName) {
                    case ("Case Number"):
                        caseNumCellRef = i;
                        break;
                    case ("Support Type"):
                        caseSupTypeRefCell = i;
                        break;
                    case ("Status"):
                        caseStatRefCell = i;
                        break;
                    case ("Severity"):
                        caseSevRefCell = i;
                        break;
                    case ("Currently Responsible"):
                        caseRespRefCell = i;
                        break;
                    case ("Case Owner"):
                        caseOwnerRefCell = i;
                        break;
                    case ("Escalated By"):
                        caseEscalatedRefCell = i;
                        break;
                    case ("Support Hotlist Level"):
                        caseHotListRefCell = i;
                        break;
                    case ("Outage Follow-Up"):
                        caseOutFolRefCell = i;
                        break;
                    case ("Age (Days)"):
                        caseAgeRefCell = i;
                        break;
                    case ("Next Case Update"):
                        caseNextUpdateDateRef = i;
                        break;
                }
            }

            for (int i = 1; i < lastRow + 1; i++) {

                caseStat = filtersheet.getRow(i).getCell(caseStatRefCell);
                String caseStatus = caseStat.getStringCellValue();

                caseSev = filtersheet.getRow(i).getCell(caseSevRefCell);
                String caseSever = caseSev.getStringCellValue();

                curResp = filtersheet.getRow(i).getCell(caseRespRefCell);
                String responsible = curResp.getStringCellValue();

                caseOwner = filtersheet.getRow(i).getCell(caseOwnerRefCell);
                String caseOwn = caseOwner.getStringCellValue();

                escCases = filtersheet.getRow(i).getCell(caseEscalatedRefCell);
                String escalatedCases = escCases.getStringCellValue();

                filtHotList = filtersheet.getRow(i).getCell(caseHotListRefCell);
                String strHotList = filtHotList.getStringCellValue();

                outfollow = filtersheet.getRow(i).getCell(caseOutFolRefCell);
                String followOut = outfollow.getStringCellValue();

                ageCase = filtersheet.getRow(i).getCell(caseAgeRefCell);
                String caseAge = ageCase.getStringCellValue();
                int ageCaseNum = Integer.parseInt(caseAge);

                caseUpdate = filtersheet.getRow(i).getCell(caseNextUpdateDateRef);
                String caseupdate = caseUpdate.getStringCellValue();

                LocalDate dateToday = LocalDate.now();
                LocalDate caseUpdateDate = null;

                if (!caseupdate.equals("NotSet")) {

                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                    caseUpdateDate = LocalDate.parse(caseupdate, formatter);
                }

                if (!strHotList.equals("NotSet") && !strHotList.equals("FALSE") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                    hotlist++;
                }

                if (caseOwn.startsWith("PS")) {
                    queuePS++;
                }

                if (caseOwn.startsWith("TS") || caseOwn.startsWith("Tech-Ops")) {
                    queueTS++;
                }

                if (followOut.equals("1") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                    outFollow++;
                }
                if (!escalatedCases.equals("NotSet") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                    escCase++;
                }

                if (caseSever.equals("Critical") && !caseStatus.equals("Pending Closure")) {
                    e1Cases++;
                }

                if (caseSever.equals("E2") && !caseStatus.equals("Pending Closure")) {
                    e2Cases++;
                }

                if (caseSever.equals("Business Critical")) {
                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        if (!caseStatus.equals("Develop Solution")) {
                            if (ageCaseNum < 15) {
                                bcDue++;
                            }
                            if (ageCaseNum > 15) {
                                misBCdue++;
                            }
                        } else {
                            BCds++;
                        }
                        bcCases++;
                    } else {
                        BCpc++;
                    }
                    if (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault"))) {
                        BCwip++;
                    }
                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        if (responsible.equals("Customer action")) {
                            custActBC++;
                        }
                        if (responsible.equals("Customer updated")) {
                            custRpdBC++;
                        }
                    }
                }

                if (caseSever.equals("Major")) {
                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        if (!caseStatus.equals("Develop Solution")) {
                            if (ageCaseNum < 30) {
                                dueMJday++;
                            }
                            if (ageCaseNum > 30) {
                                misMJdue++;
                            }
                        } else {
                            MJds++;
                        }
                    } else {
                        MJpc++;
                    }
                    if (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault"))) {
                        MJwip++;
                    }
                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        if (responsible.equals("Customer action")) {
                            custActMJ++;
                        }
                        if (responsible.equals("Customer updated")) {
                            custRpdMJ++;
                        }
                    }
                }
                if (caseSever.equals("Minor")) {
                    if (!caseStatus.equals("Develop Solution") && !caseStatus.equals("Future Availability")) {
                        if (ageCaseNum > 180) {
                            misMNdue++;
                        }
                    }
                }

                if (caseStatus.equals("Pending Closure") || caseStatus.equals("Future Availability")) {
                    inactiveCases++;
                }
                if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                    wohCases++;
                }

                if ((caseUpdateDate != null)) {
                    if (caseUpdateDate.compareTo(dateToday) == 0) {
                        updateToday++;
                    }
                    if (caseUpdateDate.compareTo(dateToday) < 0) {
                        updateMissed++;
                    }
                }

                if (caseupdate.equals("NotSet") && !caseStatus.equals("Pending Closure")) {
                    updateNull++;
                }
            }

            /* Updating the button text from digested data */
            btnHotIssues.setText(String.valueOf(hotlist));
            btnOutFollow.setText(String.valueOf(this.outFollow));
            btnEscalated.setText(String.valueOf(escCase));
            btnBCCases.setText(String.valueOf(bcCases));
            btnInactive.setText(String.valueOf(inactiveCases));
            btnBCDue.setText(String.valueOf(bcDue));
            btnBCMissed.setText(String.valueOf(misBCdue));
            btnBCWac.setText(String.valueOf(custActBC));
            btnBCupdated.setText(String.valueOf(custRpdBC));
            btnBCEngineering.setText(String.valueOf(BCds));
            btnBCINACT.setText(String.valueOf(BCpc));
            btnBCWIP.setText(String.valueOf(BCwip));
            btnMJDue.setText(String.valueOf(dueMJday));
            btnMJMissed.setText(String.valueOf(misMJdue));
            btnMJWac.setText(String.valueOf(custActMJ));
            btnMJupdated.setText(String.valueOf(custRpdMJ));
            btnMJEngineering.setText(String.valueOf(MJds));
            btnMJINACT.setText(String.valueOf(MJpc));
            btnMJWIP.setText(String.valueOf(MJwip));
            btnPSQueue.setText(String.valueOf(queuePS));
            btnTSQueue.setText(String.valueOf(queueTS));
            btnE1Cases.setText(String.valueOf(e1Cases));
            btnE2Cases.setText(String.valueOf(e2Cases));
            btnWOH.setText(String.valueOf(wohCases));
            btnUpdateToday.setText(String.valueOf(updateToday));
            btnUpdateMissed.setText(String.valueOf(updateMissed));
            btnUpdateNull.setText(String.valueOf(updateNull));
            btnMNMissed.setText(String.valueOf(misMNdue));

            /* Updating completed for overview page */

            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void surveyPage() {

        //TODO: Download the survey report and work on it


    }

    private void myCasesPage() {

        HSSFCell caseUser;
        HSSFCell myfiltHotList;
        HSSFCell myoutFollow;
        HSSFCell myescCases;
        HSSFCell mycaseSev;
        HSSFCell mycaseStat;
        HSSFCell myageCase;
        HSSFCell mycurResp;
        HSSFCell caseQueue;
        HSSFCell mycaseUpdate;


        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            myHotList = 0;
            myOutFollow = 0;
            myEscCases = 0;
            myBCCases = 0;
            myInactiveCases = 0;
            myBCDueCases = 0;
            myBCMissedCases = 0;
            myBCDSCases = 0;
            myBCInactiveCases = 0;
            myBCWIP = 0;
            myMJDueCases = 0;
            myMJMissedCases = 0;
            myMJUpdated = 0;
            myMJDSCases = 0;
            myMJWIP = 0;
            myQueuedCases = 0;
            myE1Case = 0;
            myE2Cases = 0;
            myBCupdated = 0;
            myBCWac = 0;
            myMJWAC = 0;
            myMJInactiveCases = 0;
            myWOHCases = 0;
            myUpdateToday = 0;
            myUpdateMissed = 0;
            myUpdateNull = 0;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                switch (filterColName) {
                    case ("Case Number"):
                        mycaseNumCellRef = i;
                        break;
                    case ("Support Type"):
                        mycaseSupTypeRefCell = i;
                        break;
                    case ("Status"):
                        mycaseStatRefCell = i;
                        break;
                    case ("Severity"):
                        mycaseSevRefCell = i;
                        break;
                    case ("Currently Responsible"):
                        mycaseRespRefCell = i;
                        break;
                    case ("Case Owner"):
                        mycaseOwnerRefCell = i;
                        break;
                    case ("Escalated By"):
                        mycaseEscalatedRefCell = i;
                        break;
                    case ("Support Hotlist Level"):
                        mycaseHotListRefCell = i;
                        break;
                    case ("Outage Follow-Up"):
                        mycaseOutFolRefCell = i;
                        break;
                    case ("Age (Days)"):
                        mycaseAgeRefCell = i;
                        break;
                    case ("Next Case Update"):
                        mycaseUpdateCell = i;
                        break;
                }
            }

            /* Creating Input Data Arrays from Setttings Page */

            if (!txUsers.getText().isEmpty() || !txQueues.getText().isEmpty()) {

                String userFilter = txUsers.getText();

                ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                ArrayList<String> setQueue = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                int userfiltnum = setUser.size();
                int userqueue = setQueue.size();
                ArrayList<String> mergedUsers = new ArrayList<>();

                for (int i = 0; i < userfiltnum; i++) {
                    mergedUsers.add(setUser.get(i));
                }

                for (int i = 0; i < userqueue; i++) {
                    mergedUsers.add(setQueue.get(i));

                }

                int mergedUsersCount = mergedUsers.size();

                if ((!mergedUsers.isEmpty())) {

                    for (int j = 0; j < mergedUsersCount; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            caseUser = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseuser = caseUser.getStringCellValue();

                            mycaseStat = filtersheet.getRow(i).getCell(mycaseStatRefCell);
                            String mycaseStatus = mycaseStat.getStringCellValue();

                            mycaseSev = filtersheet.getRow(i).getCell(mycaseSevRefCell);
                            String mycaseSever = mycaseSev.getStringCellValue();

                            mycurResp = filtersheet.getRow(i).getCell(mycaseRespRefCell);
                            String myresponsible = mycurResp.getStringCellValue();

                            myescCases = filtersheet.getRow(i).getCell(mycaseEscalatedRefCell);
                            String myescalatedCases = myescCases.getStringCellValue();

                            myfiltHotList = filtersheet.getRow(i).getCell(mycaseHotListRefCell);
                            String mystrFltStatus = myfiltHotList.getStringCellValue();

                            myoutFollow = filtersheet.getRow(i).getCell(mycaseOutFolRefCell);
                            String myfollowOut = myoutFollow.getStringCellValue();

                            mycaseUpdate = filtersheet.getRow(i).getCell(mycaseUpdateCell);
                            String myCaseUpdate = mycaseUpdate.getStringCellValue();

                            LocalDate dateToday = LocalDate.now();
                            LocalDate caseUpdateDate = null;

                            if (!myCaseUpdate.equals("NotSet")) {

                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                caseUpdateDate = LocalDate.parse(myCaseUpdate, formatter);
                            }

                            myageCase = filtersheet.getRow(i).getCell(mycaseAgeRefCell);
                            String mycaseAge = myageCase.getStringCellValue();
                            String myagenum = mycaseAge.replace(".0000000000", "");
                            int ageCaseNum = Integer.parseInt(myagenum);

                            if (caseuser.equals(mergedUsers.get(j))) {

                                if (!mystrFltStatus.equals("NotSet") && !mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {
                                    myHotList++;
                                }
                                if (myfollowOut.equals("1") && !mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {
                                    myOutFollow++;
                                }
                                if (!myescalatedCases.equals("NotSet") && !mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {
                                    myEscCases++;
                                }
                                if (mycaseSever.equals("Critical") && !mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {
                                    myE1Case++;
                                }
                                if (mycaseSever.equals("E2") && !mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {
                                    myE2Cases++;
                                }

                                if (mycaseSever.equals("Business Critical")) {
                                    if (!mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {

                                        myBCCases++;
                                    }

                                    if (mycaseStatus.equals("Open / Assign") || mycaseStatus.equals("Isolate Fault")) {
                                        myBCWIP++;
                                    }

                                    if (myresponsible.equals("Customer action")) {
                                        myBCWac++;
                                    }

                                    if (myresponsible.equals("Customer updated")) {
                                        myBCupdated++;
                                    }

                                    if ((mycaseStatus.equals("Open / Assign")) || (mycaseStatus.equals("Isolate Fault"))) {
                                        if (ageCaseNum <= 15) {
                                            myBCDueCases++;
                                        }
                                        if (ageCaseNum > 15) {
                                            myBCMissedCases++;
                                        }
                                    }

                                    if (mycaseStatus.equals("Develop Solution")) {
                                        myBCDSCases++;
                                    }

                                    if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                        myBCInactiveCases++;
                                    }
                                }
                                if (mycaseSever.equals("Major")) {

                                    if (mycaseStatus.equals("Develop Solution")) {
                                        myMJDSCases++;
                                    }

                                    if ((mycaseStatus.equals("Open / Assign")) || (mycaseStatus.equals("Isolate Fault"))) {
                                        if (ageCaseNum <= 30) {
                                            myMJDueCases++;
                                        }
                                        if (ageCaseNum > 30) {
                                            myMJMissedCases++;
                                        }
                                    }
                                    if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                        myMJInactiveCases++;
                                    }
                                    if (myresponsible.equals("Customer action")) {
                                        myMJWAC++;
                                    }
                                    if (myresponsible.equals("Customer updated")) {
                                        myMJUpdated++;
                                    }
                                    if (mycaseStatus.equals("Open / Assign") || (mycaseStatus.equals("Isolate Fault"))) {
                                        myMJWIP++;
                                    }
                                }
                                if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                    myInactiveCases++;
                                } else {
                                    myWOHCases++;
                                }
                                if ((caseUpdateDate != null)) {
                                    if (caseUpdateDate.compareTo(dateToday) == 0) {
                                        myUpdateToday++;
                                    }
                                    if (caseUpdateDate.compareTo(dateToday) < 0) {
                                        myUpdateMissed++;
                                    }
                                }

                                if (myCaseUpdate.equals("NotSet") && !mycaseStatus.equals("Pending Closure")) {
                                    myUpdateNull++;
                                }
                            }
                        }
                    }
                }
            }

            if (!txQueues.getText().isEmpty()) {

                ArrayList<String> setQueue = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                int queuefiltnum = setQueue.size();

                if (!setQueue.isEmpty()) {

                    for (int k = 0; k < queuefiltnum; k++) {

                        for (int l = 0; l < lastRow + 1; l++) {

                            caseQueue = workbook.getSheetAt(0).getRow(l).getCell(mycaseOwnerRefCell);
                            String casequeue = caseQueue.getStringCellValue();

                            if (casequeue.equals(setQueue.get(k))) {
                                myQueuedCases++;
                            }
                        }
                    }
                }
            }

            btnMyE1Cases.setText(String.valueOf(myE1Case));
            btnMyE2Cases.setText(String.valueOf(myE2Cases));
            btnMyOutFollow.setText(String.valueOf(myOutFollow));
            btnMyEscalated.setText(String.valueOf(myEscCases));
            btnMyBCCases.setText(String.valueOf(myBCCases));
            btnMyHotIssues.setText(String.valueOf(myHotList));
            btnMyInactive.setText(String.valueOf(myInactiveCases));
            btnMyBCWIP.setText(String.valueOf(myBCWIP));
            btnMyBCWac.setText(String.valueOf(myBCWac));
            btnMyBCupdated.setText(String.valueOf(myBCupdated));
            btnMyBCEngineering.setText(String.valueOf(myBCDSCases));
            btnMyBCINACT.setText(String.valueOf(myBCInactiveCases));
            btnMyMJWIP.setText(String.valueOf(myMJWIP));
            btnMyMJWac.setText(String.valueOf(myMJWAC));
            btnMyMJupdated.setText(String.valueOf(myMJUpdated));
            btnMyMJEngineering.setText(String.valueOf(myMJDSCases));
            btnMyMJINACT.setText(String.valueOf(myMJInactiveCases));
            btnMyBCDue.setText(String.valueOf(myBCDueCases));
            btnMyBCMissed.setText(String.valueOf(myBCMissedCases));
            btnMyMJDue.setText(String.valueOf(myMJDueCases));
            btnMyMJMissed.setText(String.valueOf(myMJMissedCases));
            btnMyQueue.setText(String.valueOf(myQueuedCases));
            btnMyWOH.setText(String.valueOf(myWOHCases));
            btnMyUpdateToday.setText(String.valueOf(myUpdateToday));
            btnMyUpdateMissed.setText(String.valueOf(myUpdateMissed));
            btnMyUpdateNull.setText(String.valueOf(myUpdateNull));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void myProductsPage() {

        HSSFCell caseUser;
        HSSFCell myfiltHotList;
        HSSFCell myoutFollow;
        HSSFCell myescCases;
        HSSFCell mycaseSev;
        HSSFCell mycaseStat;
        HSSFCell myageCase;
        HSSFCell mycurResp;
        HSSFCell caseQueue;
        HSSFCell mycaseUpdate;
        HSSFCell productName;


        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            prodHotList = 0;
            prodOutFollow = 0;
            prodEscCases = 0;
            prodBCCases = 0;
            prodInactiveCases = 0;
            prodBCDueCases = 0;
            prodBCMissedCases = 0;
            prodBCDSCases = 0;
            prodBCInactiveCases = 0;
            prodBCWIP = 0;
            prodMJDueCases = 0;
            prodMJMissedCases = 0;
            prodMJUpdated = 0;
            prodMJDSCases = 0;
            prodMJWIP = 0;
            prodQueuedCases = 0;
            prodE1Case = 0;
            prodE2Cases = 0;
            prodBCupdated = 0;
            prodBCWac = 0;
            prodMJWAC = 0;
            prodMJInactiveCases = 0;
            prodWOHCases = 0;
            prodUpdateToday = 0;
            prodUpdateMissed = 0;
            prodUpdateNull = 0;
            prodQueuePS = 0;
            prodQueueTS = 0;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                switch (filterColName) {
                    case ("Case Number"):
                        caseNumCellRef = i;
                        break;
                    case ("Support Type"):
                        caseSupTypeRefCell = i;
                        break;
                    case ("Status"):
                        caseStatRefCell = i;
                        break;
                    case ("Severity"):
                        caseSevRefCell = i;
                        break;
                    case ("Currently Responsible"):
                        caseRespRefCell = i;
                        break;
                    case ("Case Owner"):
                        caseOwnerRefCell = i;
                        break;
                    case ("Escalated By"):
                        caseEscalatedRefCell = i;
                        break;
                    case ("Support Hotlist Level"):
                        caseHotListRefCell = i;
                        break;
                    case ("Outage Follow-Up"):
                        caseOutFolRefCell = i;
                        break;
                    case ("Age (Days)"):
                        caseAgeRefCell = i;
                        break;
                    case ("Next Case Update"):
                        caseNextUpdateDateRef = i;
                        break;
                    case ("Support Product"):
                        caseProductRef = i;
                        break;
                }
            }

            /* Creating Input Data Arrays from Setttings Page */

            if (!txProducts.getText().isEmpty()) {

                ArrayList<String> setProducts = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));
                int productSettingsNum = setProducts.size();

                if ((!setProducts.isEmpty())) {

                    for (int j = 0; j < productSettingsNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            caseUser = filtersheet.getRow(i).getCell(caseOwnerRefCell);
                            String caseuser = caseUser.getStringCellValue();

                            mycaseStat = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String mycaseStatus = mycaseStat.getStringCellValue();

                            mycaseSev = filtersheet.getRow(i).getCell(caseSevRefCell);
                            String mycaseSever = mycaseSev.getStringCellValue();

                            mycurResp = filtersheet.getRow(i).getCell(caseRespRefCell);
                            String myresponsible = mycurResp.getStringCellValue();

                            myescCases = filtersheet.getRow(i).getCell(caseEscalatedRefCell);
                            String myescalatedCases = myescCases.getStringCellValue();

                            myfiltHotList = filtersheet.getRow(i).getCell(caseHotListRefCell);
                            String mystrFltStatus = myfiltHotList.getStringCellValue();

                            myoutFollow = filtersheet.getRow(i).getCell(caseOutFolRefCell);
                            String myfollowOut = myoutFollow.getStringCellValue();

                            productName = filtersheet.getRow(i).getCell(caseProductRef);
                            String productCellStr = productName.getStringCellValue();

                            mycaseUpdate = filtersheet.getRow(i).getCell(caseNextUpdateDateRef);
                            String myCaseUpdate = mycaseUpdate.getStringCellValue();

                            LocalDate dateToday = LocalDate.now();
                            LocalDate caseUpdateDate = null;

                            if (!myCaseUpdate.equals("NotSet")) {

                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                caseUpdateDate = LocalDate.parse(myCaseUpdate, formatter);
                            }

                            myageCase = filtersheet.getRow(i).getCell(caseAgeRefCell);
                            String mycaseAge = myageCase.getStringCellValue();
                            String myagenum = mycaseAge.replace(".0000000000", "");
                            int ageCaseNum = Integer.parseInt(myagenum);

                            if (productCellStr.equals(setProducts.get(j))) {

                                if (!mystrFltStatus.equals("NotSet") && !mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {
                                    prodHotList++;
                                }
                                if (myfollowOut.equals("1") && !mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {
                                    prodOutFollow++;
                                }
                                if (!myescalatedCases.equals("NotSet") && !mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {
                                    prodEscCases++;
                                }
                                if (mycaseSever.equals("Critical") && !mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {
                                    prodE1Case++;
                                }
                                if (mycaseSever.equals("E2") && !mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {
                                    prodE2Cases++;
                                }

                                if (mycaseSever.equals("Business Critical")) {
                                    if (!mycaseStatus.equals("Pending Closure") && !mycaseStatus.equals("Future Availability")) {

                                        prodBCCases++;
                                    }

                                    if (mycaseStatus.equals("Open / Assign") || mycaseStatus.equals("Isolate Fault")) {
                                        prodBCWIP++;
                                    }

                                    if (myresponsible.equals("Customer action")) {
                                        prodBCWac++;
                                    }

                                    if (myresponsible.equals("Customer updated")) {
                                        prodBCupdated++;
                                    }

                                    if (!mycaseStatus.equals("Develop Solution") || !mycaseStatus.equals("Future Availability")) {
                                        if (ageCaseNum < 15) {
                                            prodBCDueCases++;
                                        } else {
                                            prodBCMissedCases++;
                                        }
                                    }

                                    if (mycaseStatus.equals("Develop Solution")) {
                                        prodBCDSCases++;
                                    }

                                    if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                        prodBCInactiveCases++;
                                    }
                                }
                                if (mycaseSever.equals("Major")) {

                                    if (mycaseStatus.equals("Develop Solution")) {
                                        prodMJDSCases++;
                                    }

                                    if (!mycaseStatus.equals("Develop Solution") || !mycaseStatus.equals("Future Availability")) {
                                        if (ageCaseNum < 30) {
                                            prodMJDueCases++;
                                        } else {
                                            prodMJMissedCases++;
                                        }
                                    }
                                    if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                        prodMJInactiveCases++;
                                    }
                                    if (myresponsible.equals("Customer action")) {
                                        prodMJWAC++;
                                    }
                                    if (myresponsible.equals("Customer updated")) {
                                        prodMJUpdated++;
                                    }
                                    if (mycaseStatus.equals("Open / Assign") || (mycaseStatus.equals("Isolate Fault"))) {
                                        prodMJWIP++;
                                    }
                                }
                                if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                    prodInactiveCases++;
                                } else {
                                    prodWOHCases++;
                                }
                                if ((caseUpdateDate != null)) {
                                    if (caseUpdateDate.compareTo(dateToday) == 0) {
                                        prodUpdateToday++;
                                    }
                                    if (caseUpdateDate.compareTo(dateToday) < 0) {
                                        prodUpdateMissed++;
                                    }
                                }

                                if (myCaseUpdate.equals("NotSet") && !mycaseStatus.equals("Pending Closure")) {
                                    prodUpdateNull++;
                                }
                                if (caseuser.startsWith("PS ")) {
                                    prodQueuePS++;

                                }
                                if (caseuser.startsWith("TS ") || caseuser.startsWith("Tech-Ops")) {
                                    prodQueueTS++;
                                }

                            }
                        }
                    }
                }
            }

            if (!txQueues.getText().isEmpty()) {

                ArrayList<String> setQueue = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                int queuefiltnum = setQueue.size();

                if (!setQueue.isEmpty()) {

                    for (int k = 0; k < queuefiltnum; k++) {

                        for (int l = 0; l < lastRow + 1; l++) {

                            caseQueue = workbook.getSheetAt(0).getRow(l).getCell(caseOwnerRefCell);
                            String casequeue = caseQueue.getStringCellValue();

                            if (casequeue.equals(setQueue.get(k))) {
                                prodQueuedCases++;
                            }
                        }
                    }
                }
            }

            btnE1Prod.setText(String.valueOf(prodE1Case));
            btnE2Prod.setText(String.valueOf(prodE2Cases));
            btnOutFollowProd.setText(String.valueOf(prodOutFollow));
            btnEscalatedProd.setText(String.valueOf(prodEscCases));
            btnBCProd.setText(String.valueOf(prodBCCases));
            btnHotIssuesProd.setText(String.valueOf(prodHotList));
            btnInactiveProd.setText(String.valueOf(prodInactiveCases));
            btnBCWIPProd.setText(String.valueOf(prodBCWIP));
            btnBCWacProd.setText(String.valueOf(prodBCWac));
            btnBCupdatedProd.setText(String.valueOf(prodBCupdated));
            btnBCEngineeringProd.setText(String.valueOf(prodBCDSCases));
            btnBCINACTProd.setText(String.valueOf(prodBCInactiveCases));
            btnMJWIPProd.setText(String.valueOf(prodMJWIP));
            btnMJWacProd.setText(String.valueOf(prodMJWAC));
            btnMJupdatedProd.setText(String.valueOf(prodMJUpdated));
            btnMJEngineeringProd.setText(String.valueOf(prodMJDSCases));
            btnMJINACTProd.setText(String.valueOf(prodMJInactiveCases));
            btnBCDueProd.setText(String.valueOf(prodBCDueCases));
            btnBCMissedProd.setText(String.valueOf(prodBCMissedCases));
            btnMJDueProd.setText(String.valueOf(prodMJDueCases));
            btnMJMissedProd.setText(String.valueOf(prodMJMissedCases));
            btnPSQueueProd.setText(String.valueOf(prodQueuePS));
            btnTSQueueProd.setText(String.valueOf(prodQueueTS));
            btnWOHProd.setText(String.valueOf(prodWOHCases));
            //btnMyUpdateToday.setText(String.valueOf(myUpdateToday));
            //btnMyUpdateMissed.setText(String.valueOf(myUpdateMissed));
            //btnMyUpdateNull.setText(String.valueOf(myUpdateNull));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @FXML
    void handleMouseOver(MouseEvent event) {
        if (event.getSource() == btnHome) {

            Tooltip homeTooltip = new Tooltip();
            homeTooltip.setText("General View\n" +
                    "Not Closed Cases");
            btnHome.setTooltip(homeTooltip);
        }

        if (event.getSource() == btnCases) {

            Tooltip casesTooltip = new Tooltip();
            casesTooltip.setText("Personalized View");
            btnCases.setTooltip(casesTooltip);

        }

        if (event.getSource() == btnCustomers) {

            Tooltip customersTooltip = new Tooltip();
            customersTooltip.setText("Customer Based\n" +
                    "Case View");
            btnCustomers.setTooltip(customersTooltip);

        }

        if (event.getSource() == btnSurvey) {

            Tooltip surveyTooltip = new Tooltip();
            surveyTooltip.setText("VERY SOON");
            btnSurvey.setTooltip(surveyTooltip);

        }

        if (event.getSource() == btnProjects) {

            Tooltip projectTooltip = new Tooltip();
            projectTooltip.setText("VERY SOON");
            btnProjects.setTooltip(projectTooltip);

        }

        if (event.getSource() == btnSettings) {

            Tooltip settingsTooltip = new Tooltip();
            settingsTooltip.setText("Personalize\n" +
                    "Your Querries");
            btnSettings.setTooltip(settingsTooltip);

        }

        if (event.getSource() == btnLoadData) {

            Tooltip loadTooltip = new Tooltip();
            loadTooltip.setText("Connect to SFDC and \n" +
                    "gather recent data");
            btnLoadData.setTooltip(loadTooltip);

        }

        if (event.getSource() == txUsers) {

            Tooltip userTextBoxTip = new Tooltip();
            userTextBoxTip.setText("Please prompt user names as\n" +
                    "provisioned in Salesforce!");
            txUsers.setTooltip(userTextBoxTip);
        }
    }

    private void initTableView(TableView<CaseTableView> table) {

        //tableCases.getStyleClass().add("table-view");

        if (table == tableCases) {

            NumberCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseNumber"));
            SeverityCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseSeverity"));
            StatusCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseStatus"));
            OwnerCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseOwner"));
            ResponsibleCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseResponsible"));
            AgeCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseAge"));
            NextUpdateCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, LocalDate>("nextCaseUpdate"));
            EscalatedByCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseEscalatedBy"));
            HotListCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseHotList"));
            OutFollowCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseOutFollow"));
            SupportTypeCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseSupportType"));
            ProductCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseProduct"));
            SubjectCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseSubject"));
            AccountCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseAccount"));
            RegionCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseRegion"));
            SecurityCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseSecurity"));
            DateTimeOpenedCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseDateTimeOpened"));
        }
        if (table == tableCustomers) {

            NumberColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseNumber"));
            SeverityColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseSeverity"));
            StatusColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseStatus"));
            OwnerColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseOwner"));
            ResponsibleColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseResponsible"));
            AgeColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseAge"));
            NextUpdateColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, LocalDate>("nextCaseUpdate"));
            EscalatedByColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseEscalatedBy"));
            HotListColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseHotList"));
            OutFollowColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseOutFollow"));
            SupportTypeColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseSupportType"));
            ProductColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseProduct"));
            SubjectColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseSubject"));
            AccountColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseAccount"));
            RegionColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseRegion"));
            SecurityColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseSecurity"));
            DateTimeOpenedColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseDateTimeOpened"));

        }
    }

    private void copyCaseNumberToClipboard(TableView<CaseTableView> tableCases) {

        TablePosition tablePosition = (TablePosition) tableCases.getSelectionModel().getSelectedCells().get(0);
        int row = tablePosition.getRow();
        CaseTableView caseview = (CaseTableView) tableCases.getItems().get(row);
        TableColumn tableColumn = tablePosition.getTableColumn();
        String data1 = caseview.getCaseNumber();
        ClipboardContent content = new ClipboardContent();
        content.putString(data1);
        Clipboard.getSystemClipboard().setContent(content);

    }

    private String getCaseNumber(TableView<CaseTableView> tableCases, String caseNumber){

        TablePosition tablePosition = (TablePosition) tableCases.getSelectionModel().getSelectedCells().get(0);
        int row = tablePosition.getRow();
        CaseTableView caseview = (CaseTableView) tableCases.getItems().get(row);
        TableColumn tableColumn = tablePosition.getTableColumn();
        caseNumber = caseview.getCaseNumber();
        return caseNumber;
    }


    private void exportExcelAction(TableView<CaseTableView> table) {

        try {
            FileChooser fileChooser = new FileChooser();
            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("XLS files (*.xls)", "*.xls");
            fileChooser.setInitialDirectory(new File(System.getProperty("user.home") + "\\Desktop"));

            fileChooser.getExtensionFilters().add(extFilter);

            Stage primaryStage = new Stage();

            File file = fileChooser.showSaveDialog(primaryStage);
            primaryStage.show();

            if (file != null) {

                extractToExcel(table, "testData", file);
            }
            primaryStage.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void customerViewPage() {

        HSSFCell account;
        HSSFCell severity;
        HSSFCell age;
        HSSFCell nextUpdate;
        HSSFCell hotIssue;
        HSSFCell escalated;
        HSSFCell status;
        HSSFCell outFollow;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                switch (filterColName) {
                    case ("Status"):
                        caseStatRefCell = i;
                        break;
                    case ("Severity"):
                        caseSevRefCell = i;
                        break;
                    case ("Escalated By"):
                        caseEscalatedRefCell = i;
                        break;
                    case ("Support Hotlist Level"):
                        caseHotListRefCell = i;
                        break;
                    case ("Outage Follow-Up"):
                        caseOutFolRefCell = i;
                        break;
                    case ("Age (Days)"):
                        caseAgeRefCell = i;
                        break;
                    case ("Next Case Update"):
                        caseNextUpdateDateRef = i;
                        break;
                    case ("Account Name"):
                        caseAccountRef = i;
                }
            }

            if (!customerText.getText().isEmpty()) {

                ArrayList<String> setCustomerList = new ArrayList<>(Arrays.asList(customerText.getText().split(",\\s*")));

                //ArrayList<String> setCustomerAsItis = new ArrayList<>(Arrays.asList(customerText.getText().split(",\\s*")));
                //ArrayList<String> setCustomerCap = new ArrayList();

                int customerNum = setCustomerList.size();

                customerE1 = 0;
                customerE2 = 0;
                customerOutFol = 0;
                customerHot = 0;
                customerEsc = 0;
                customerBC = 0;
                customerWoh = 0;

                /*for (int i = 0; i < customerNum; i++) {

                    Pattern pattern = Pattern.compile("\\b([a-z])([\\w]*)");
                    Matcher matcher = pattern.matcher(setCustomerList.get(i));
                    StringBuffer buffer = new StringBuffer();
                    while (matcher.find()) {
                        matcher.appendReplacement(buffer, matcher.group(1).toUpperCase() + matcher.group(2));
                    }
                    String capitalized = matcher.appendTail(buffer).toString();
                    setCustomerCap.add(capitalized);
                }

                int setcust2num = setCustomerCap.size();*/


                if ((!setCustomerList.isEmpty())) {

                    for (int j = 0; j < customerNum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            status = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = status.getStringCellValue();

                            severity = filtersheet.getRow(i).getCell(caseSevRefCell);
                            String caseSeverity = severity.getStringCellValue();

                            escalated = filtersheet.getRow(i).getCell(caseEscalatedRefCell);
                            String escalatedCases = escalated.getStringCellValue();

                            hotIssue = filtersheet.getRow(i).getCell(caseHotListRefCell);
                            String hotIssueCases = hotIssue.getStringCellValue();

                            outFollow = filtersheet.getRow(i).getCell(caseOutFolRefCell);
                            String outFollowCases = outFollow.getStringCellValue();

                            account = filtersheet.getRow(i).getCell(caseAccountRef);
                            String accountName = account.getStringCellValue();

                            nextUpdate = filtersheet.getRow(i).getCell(caseNextUpdateDateRef);
                            String nextUpdateCase = nextUpdate.getStringCellValue();

                            if (accountName.equals(setCustomerList.get(j))) {

                                if (caseSeverity.equals("Critical") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    customerE1++;
                                }
                                if (caseSeverity.equals("E2") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    customerE2++;
                                }
                                if (outFollowCases.equals("1") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    customerOutFol++;
                                }
                                if (!hotIssueCases.equals("NotSet") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    customerHot++;
                                }
                                if (!escalatedCases.equals("NotSet") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    customerEsc++;
                                }
                                if (caseSeverity.equals("Business Critical") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    customerBC++;
                                }
                                if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    customerWoh++;
                                }
                            }
                        }
                    }
                }
            }

            btnCustomerCritical.setText(String.valueOf(customerE1));
            btnCustomerE2.setText(String.valueOf(customerE2));
            btnCustomerHotIssues.setText(String.valueOf(customerOutFol));
            btnCustomerEscalated.setText(String.valueOf(customerEsc));
            btnCustomerHotIssues.setText(String.valueOf(customerHot));
            btnCustomerBC.setText(String.valueOf(customerBC));
            btnCustomerActWOH.setText(String.valueOf(customerWoh));
            btnCustomerOutFollow.setText(String.valueOf(customerOutFol));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void alertUser() {
        Alert alert = new Alert(Alert.AlertType.WARNING);
        ((Stage) alert.getDialogPane().getScene().getWindow()).getIcons().add(new Image("home/image/rbbicon.PNG"));
        alert.setTitle("RBBN CASE MANAGEMENT TOOL WARNING:");
        alert.setHeaderText(null);
        alert.setContentText("NO RECORD FOUND..." + "\n" + "\n" + "PLEASE RELOAD DATA FOR RECENT DATA!" + "\n" + "\n" + "IF NOT ALREADY, PLEASE LOGIN!");
        alert.showAndWait();
    }

    private void userSelectArray() {

        HSSFCell userCell;

        tableUsers.setVisible(true);
        userCol.setCellValueFactory(new PropertyValueFactory<UserTableView, String>("userName"));
        userSelectedCol.setCellValueFactory(new PropertyValueFactory<UserTableView, String>("userName"));

        try {

            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_user_prod.xls")));
            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Case Owner")) {
                    caseOwnerRefCell = i;
                }
            }

            ArrayList<String> userArray = new ArrayList<>();

            for (int i = 1; i < lastRow; i++) {

                userCell = filtersheet.getRow(i).getCell(caseOwnerRefCell);
                String userName = userCell.getStringCellValue();

                if (!userName.startsWith("PS ") && !userName.startsWith("TS ") && !userName.startsWith("Tech-Ops ")) {
                    userArray.add(userName);
                }

            }

            userArray = (ArrayList) userArray.stream().distinct().collect(Collectors.toList());

            ObservableList<UserTableView> usrList = FXCollections.observableArrayList();

            int arraysize = userArray.size();

            for (int i = 0; i < arraysize; i++) {

                usrList.add(new UserTableView(userArray.get(i)));
            }

            FilteredList<UserTableView> filteredUsers = new FilteredList((ObservableList) usrList, p -> true);
            txtUserSelect.textProperty().addListener((observable, oldValue, newValue) -> {
                filteredUsers.setPredicate(userTableView -> {

                    if (newValue == null || newValue.isEmpty()) {
                        return true;
                    }

                    String lowerCaseCustomerName = newValue.toLowerCase();

                    if (userTableView.getUserName().toLowerCase().contains(lowerCaseCustomerName)) {
                        return true;
                    }
                    return false;
                });
            });


            SortedList<UserTableView> sortedUser = new SortedList<>(filteredUsers);
            sortedUser.comparatorProperty().bind(tableUsers.comparatorProperty());

            tableUsers.setItems(filteredUsers);
            tableUsers.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
            tableUsers.getSelectionModel().setCellSelectionEnabled(true);

            tableUsers.getFocusModel().focusedCellProperty().addListener((obs, newVal, oldVal) -> {

                tableUsers.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        if (event.getClickCount() > 1) {
                            try {

                                if (tableUsers.getSelectionModel().getSelectedItem() != null) {
                                    UserTableView selectedUsr = tableUsers.getSelectionModel().getSelectedItem();
                                    //filteredAccounts.add(selectedAcc.getAccountName());
                                    tableUsersSelected.getItems().add(selectedUsr);
                                }

                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }
                    }
                });

            });

            tableUsersSelected.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    if (event.getClickCount() > 1) {
                        try {

                            if (tableUsersSelected.getSelectionModel().getSelectedCells() != null) {
                                UserTableView selectedCust = tableUsersSelected.getSelectionModel().getSelectedItem();
                                tableUsersSelected.getItems().remove(selectedCust);
                            }

                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }

                }
            });

            btnUsersUpdate.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    int selected = 0;
                    usersFiltered = new ArrayList<>();

                    try {

                        selected = tableUsersSelected.getItems().size();

                        for (int i = 0; i < selected; i++) {

                            UserTableView addUsr = tableUsersSelected.getItems().get(i);
                            usersFiltered.add(addUsr.getUserName());

                        }

                        usersFiltered = (ArrayList) usersFiltered.stream().distinct().collect(Collectors.toList());

                    } catch (Exception e) {
                        e.printStackTrace();
                    }

                    txUsers.setText(usersFiltered.toString().replace("[", "").replace("]", ""));
                    pnUsersSelect.setVisible(false);
                }
            });

            btnUserSelectClear.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    tableUsersSelected.getItems().clear();
                }
            });

            btnUserSelectClose.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    pnUsersSelect.setVisible(false);
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void productSelectArray() {

        HSSFCell prodCell;

        tableProducts.setVisible(true);
        productCol.setCellValueFactory(new PropertyValueFactory<ProductTableView, String>("productName"));
        productColSelected.setCellValueFactory(new PropertyValueFactory<ProductTableView, String>("productName"));

        try {

            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_user_prod.xls")));
            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Support Product")) {
                    caseProductRef = i;
                }
            }

            ArrayList<String> prodArray = new ArrayList<>();

            for (int i = 1; i < lastRow; i++) {

                prodCell = filtersheet.getRow(i).getCell(caseProductRef);
                String productName = prodCell.getStringCellValue();
                prodArray.add(productName);
            }

            prodArray = (ArrayList) prodArray.stream().distinct().collect(Collectors.toList());

            ObservableList<ProductTableView> prodList = FXCollections.observableArrayList();

            int arraysize = prodArray.size();

            for (int i = 0; i < arraysize; i++) {

                prodList.add(new ProductTableView(prodArray.get(i)));
            }

            FilteredList<ProductTableView> filteredProducts = new FilteredList((ObservableList) prodList, p -> true);
            txtProductSelect.textProperty().addListener((observable, oldValue, newValue) -> {
                filteredProducts.setPredicate(productTableView -> {

                    if (newValue == null || newValue.isEmpty()) {
                        return true;
                    }

                    String lowerCaseCustomerName = newValue.toLowerCase();

                    if (productTableView.getProductName().toLowerCase().contains(lowerCaseCustomerName)) {
                        return true;
                    }
                    return false;
                });
            });

            SortedList<ProductTableView> sortedProducts = new SortedList<>(filteredProducts);
            sortedProducts.comparatorProperty().bind(tableProducts.comparatorProperty());

            tableProducts.setItems(filteredProducts);
            tableProducts.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
            tableProducts.getSelectionModel().setCellSelectionEnabled(true);

            tableProducts.getFocusModel().focusedCellProperty().addListener((obs, newVal, oldVal) -> {

                tableProducts.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        if (event.getClickCount() > 1) {
                            try {

                                if (tableProducts.getSelectionModel().getSelectedItem() != null) {
                                    ProductTableView selectedProduct = tableProducts.getSelectionModel().getSelectedItem();
                                    //filteredAccounts.add(selectedAcc.getAccountName());
                                    tableProductsSelected.getItems().add(selectedProduct);
                                }

                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }
                    }
                });

            });

            tableProductsSelected.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    if (event.getClickCount() > 1) {
                        try {

                            if (tableProductsSelected.getSelectionModel().getSelectedCells() != null) {
                                ProductTableView selectedCust = tableProductsSelected.getSelectionModel().getSelectedItem();
                                tableProductsSelected.getItems().remove(selectedCust);
                            }

                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }

                }
            });

            btnProductUpdate.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    int selected = 0;
                    productsFiltered = new ArrayList<>();

                    try {

                        selected = tableProductsSelected.getItems().size();

                        for (int i = 0; i < selected; i++) {

                            ProductTableView addUsr = tableProductsSelected.getItems().get(i);
                            productsFiltered.add(addUsr.getProductName());

                        }

                        productsFiltered = (ArrayList) productsFiltered.stream().distinct().collect(Collectors.toList());

                    } catch (Exception e) {
                        e.printStackTrace();
                    }

                    txProducts.setText(productsFiltered.toString().replace("[", "").replace("]", ""));
                    pnProductSelect.setVisible(false);
                }
            });

            btnProductSelectClear.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    tableProductsSelected.getItems().clear();
                }
            });

            btnProductSelectClose.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    pnProductSelect.setVisible(false);
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void queueSelectArray() {

        tableQueue.setVisible(true);
        queueCol.setCellValueFactory(new PropertyValueFactory<QueueTableView, String>("queueName"));
        queueColSelected.setCellValueFactory(new PropertyValueFactory<QueueTableView, String>("queueName"));

        int arraySize = queueArray.size();

        ObservableList<QueueTableView> queList = FXCollections.observableArrayList();


        for (int i = 0; i < arraySize; i++) {

            queList.add(new QueueTableView(queueArray.get(i)));
        }

        FilteredList<QueueTableView> filteredQueues = new FilteredList((ObservableList) queList, p -> true);
        txtQueueSelect.textProperty().addListener((observable, oldValue, newValue) -> {
            filteredQueues.setPredicate(queueTableView -> {

                if (newValue == null || newValue.isEmpty()) {
                    return true;
                }

                String lowerCaseCustomerName = newValue.toLowerCase();

                if (queueTableView.getQueueName().toLowerCase().contains(lowerCaseCustomerName)) {
                    return true;
                }
                return false;
            });
        });

        SortedList<QueueTableView> sortedQueues = new SortedList<>(filteredQueues);
        sortedQueues.comparatorProperty().bind(tableQueue.comparatorProperty());

        tableQueue.setItems(filteredQueues);
        tableQueue.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        tableQueue.getSelectionModel().setCellSelectionEnabled(true);

        tableQueue.getFocusModel().focusedCellProperty().addListener((obs, newVal, oldVal) -> {

            tableQueue.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    if (event.getClickCount() > 1) {
                        try {

                            if (tableQueue.getSelectionModel().getSelectedItem() != null) {
                                QueueTableView selectedQue = tableQueue.getSelectionModel().getSelectedItem();
                                //filteredAccounts.add(selectedAcc.getAccountName());
                                tableQueueSelected.getItems().add(selectedQue);
                            }

                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                }
            });

        });

        tableQueueSelected.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                if (event.getClickCount() > 1) {
                    try {

                        if (tableQueueSelected.getSelectionModel().getSelectedCells() != null) {
                            QueueTableView selectedQueue = tableQueueSelected.getSelectionModel().getSelectedItem();
                            tableQueueSelected.getItems().remove(selectedQueue);
                        }

                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }

            }
        });

        btnQueueUpdate.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                int selected = 0;
                queuesFiltered = new ArrayList<>();

                try {

                    selected = tableQueueSelected.getItems().size();

                    for (int i = 0; i < selected; i++) {

                        QueueTableView addQue = tableQueueSelected.getItems().get(i);
                        queuesFiltered.add(addQue.getQueueName());

                    }

                    queuesFiltered = (ArrayList) queuesFiltered.stream().distinct().collect(Collectors.toList());

                } catch (Exception e) {
                    e.printStackTrace();
                }

                txQueues.setText(queuesFiltered.toString().replace("[", "").replace("]", ""));
                pnQueueSelect.setVisible(false);
            }
        });

        btnQueueSelectClear.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                tableQueueSelected.getItems().clear();

            }
        });

        btnQueueSelectClose.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                pnQueueSelect.setVisible(false);
            }
        });
    }


    public void accountArray() {

        HSSFCell accountCell;

        tableAccounts.setVisible(true);
        customerCol.setCellValueFactory(new PropertyValueFactory<AccountTableView, String>("accountName"));
        customerSelectedCol.setCellValueFactory(new PropertyValueFactory<AccountTableView, String>("accountName"));

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
            }

            ArrayList<String> accountArray = new ArrayList<>();

            for (int i = 1; i < lastRow; i++) {

                accountCell = filtersheet.getRow(i).getCell(caseAccountRef);
                String accountName = accountCell.getStringCellValue();
                accountArray.add(accountName);
            }

            accountArray = (ArrayList) accountArray.stream().distinct().collect(Collectors.toList());

            ObservableList<AccountTableView> actList = FXCollections.observableArrayList();

            int arraysize = accountArray.size();

            for (int i = 0; i < arraysize; i++) {

                actList.add(new AccountTableView(accountArray.get(i)));
            }

            FilteredList<AccountTableView> filteredCustomers = new FilteredList((ObservableList) actList, p -> true);
            txtFilterAccounts.textProperty().addListener((observable, oldValue, newValue) -> {
                filteredCustomers.setPredicate(accountTableView -> {

                    if (newValue == null || newValue.isEmpty()) {
                        return true;
                    }

                    String lowerCaseCustomerName = newValue.toLowerCase();

                    if (accountTableView.getAccountName().toLowerCase().contains(lowerCaseCustomerName)) {
                        return true;
                    }
                    return false;
                });
            });

            SortedList<AccountTableView> sortedCustomer = new SortedList<>(filteredCustomers);
            sortedCustomer.comparatorProperty().bind(tableAccounts.comparatorProperty());

            tableAccounts.setItems(filteredCustomers);
            tableAccounts.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
            tableAccounts.getSelectionModel().setCellSelectionEnabled(true);

            tableAccounts.getFocusModel().focusedCellProperty().addListener((obs, newVal, oldVal) -> {

                tableAccounts.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        if (event.getClickCount() > 1) {
                            try {

                                if (tableAccounts.getSelectionModel().getSelectedItem() != null) {
                                    AccountTableView selectedAcc = tableAccounts.getSelectionModel().getSelectedItem();
                                    //filteredAccounts.add(selectedAcc.getAccountName());
                                    tableAccountsSelected.getItems().add(selectedAcc);
                                }

                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }
                    }
                });

            });

            tableAccountsSelected.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    if (event.getClickCount() > 1) {
                        try {

                            if (tableAccountsSelected.getSelectionModel().getSelectedCells() != null) {
                                AccountTableView selectedCust = tableAccountsSelected.getSelectionModel().getSelectedItem();
                                tableAccountsSelected.getItems().remove(selectedCust);
                            }

                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }

                }
            });

            btnFilterAccountUpdate.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    btnCustomerLoad.setVisible(true);
                    int selected = 0;
                    filteredAccounts = new ArrayList<>();

                    try {

                        selected = tableAccountsSelected.getItems().size();

                        for (int i = 0; i < selected; i++) {

                            AccountTableView addCust = tableAccountsSelected.getItems().get(i);
                            filteredAccounts.add(addCust.getAccountName());

                        }

                        filteredAccounts = (ArrayList) filteredAccounts.stream().distinct().collect(Collectors.toList());

                    } catch (Exception e) {
                        e.printStackTrace();
                    }

                    customerText.setText(filteredAccounts.toString().replace("[", "").replace("]", ""));
                    pnAccountSelect.setVisible(false);
                }
            });

            btnFilterAccountClose.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    pnAccountSelect.setVisible(false);
                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void initCustomerNumbers() {

        btnCustomerCritical.setText("0");
        btnCustomerE2.setText("0");
        btnCustomerHotIssues.setText("0");
        btnCustomerEscalated.setText("0");
        btnCustomerHotIssues.setText("0");
        btnCustomerBC.setText("0");
        btnCustomerActWOH.setText("0");
        btnCustomerOutFollow.setText("0");
    }

    @FXML
    private void handleMouseClicked(MouseEvent event) {
        if (event.getSource() == customerText) {
            pnAccountSelect.setVisible(true);
            pnAccountSelect.toFront();
            accountArray();
            txtFilterAccounts.requestFocus();
        }
        if (event.getSource() == apnCustomers) {
            pnAccountSelect.setVisible(false);
        }

        if (event.getSource() == txUsers) {
            pnUsersSelect.setVisible(true);
            pnProductSelect.setVisible(false);
            pnQueueSelect.setVisible(false);
            userSelectArray();
            txtUserSelect.requestFocus();
        }

        if (event.getSource() == apnSettings) {
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnQueueSelect.setVisible(false);

        }
        if (event.getSource() == txProducts) {
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(true);
            pnQueueSelect.setVisible(false);
            productSelectArray();
            txtProductSelect.requestFocus();
        }

        if (event.getSource() == txQueues) {
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnQueueSelect.setVisible(true);
            txtQueueSelect.requestFocus();
        }
        if (event.getSource() == btnUsersClear) {
            pnUsersSelect.setVisible(false);
            txUsers.clear();
        }
        if (event.getSource() == btnProductsClear) {
            pnProductSelect.setVisible(false);
            txProducts.clear();
            txtUserSelect.clear();
        }
        if (event.getSource() == btnQueueClear) {
            pnQueueSelect.setVisible(false);
            txQueues.clear();
            txtUserSelect.clear();
        }
        if (event.getSource() == btnAccountClear) {

            btnCustomerLoad.setVisible(false);
            filteredAccounts.clear();
            tableAccountsSelected.getItems().clear();
            customerText.setText("");
            tableCustomers.setVisible(false);
            initCustomerNumbers();
            pnAccountSelect.setVisible(false);
        }
        if (event.getSource() == txQueues) {
            queueSelectArray();
        }
        if (event.getSource() == btnClearAll) {
            txUsers.clear();
            txProducts.clear();
            txQueues.clear();
        }
        if (event.getSource() == btnSaveDefault) {
            String userFilter = txUsers.getText();
            String queueFilter = txQueues.getText();
            String productFilter = txProducts.getText();

            writeDefaultSettingsToFile(userFilter, queueFilter, productFilter);
        }

        if (event.getSource() == btnLoadDefault) {

            readDefaultSettingFiles();
        }

        if(event.getSource() == btnInfo){

            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            ((Stage) alert.getDialogPane().getScene().getWindow()).getIcons().add(new Image("home/image/rbbicon.PNG"));
            alert.setTitle("RBBN CASE MANAGEMENT TOOL");
            alert.setHeaderText(null);
            alert.setContentText("Designed by:" + "\n" + "\n" + "Ali Alper Simsek & Vehbi Benli" + "\n" + "\n" +
                    "For any issues please inform below:" + "\n" +
                    "asimsek@rbbn.com" + "\n" +
                    "vbenli@rbbn.com");
            alert.showAndWait();
        }
    }

    public void setqueueArray() {

        queueArray = new ArrayList<>();

        queueArray.add(0, "Kandy NOC");
        queueArray.add(1, "KBS Onboarding");
        queueArray.add(2, "KBS Operations");
        queueArray.add(3, "KBS Support");
        queueArray.add(4, "PS A2 Call Processing");
        queueArray.add(5, "PS A2 Gateways");
        queueArray.add(6, "PS A2 GENCOM");
        queueArray.add(7, "PS A2 IMM");
        queueArray.add(8, "PS A2 OAM");
        queueArray.add(9, "PS A2 WAM");
        queueArray.add(10, "PS A6");
        queueArray.add(11, "PS Billing");
        queueArray.add(12, "PS C3");
        queueArray.add(13, "PS CBM SDM");
        queueArray.add(14, "PS CCA SST SAM21 Platform");
        queueArray.add(15, "PS CICM");
        queueArray.add(16, "PS CM9520");
        queueArray.add(17, "PS Converged Intelligent Messaging (CIM)");
        queueArray.add(18, "PS CoreBase SW");
        queueArray.add(19, "PS CoreHardware");
        queueArray.add(20, "PS CPaaS");
        queueArray.add(21, "PS CSLAN8600");
        queueArray.add(22, "PS DMS SS7");
        queueArray.add(23, "PS EMT");
        queueArray.add(24, "PS G5");
        queueArray.add(25, "PS Gateways");
        queueArray.add(26, "PS GENiUS");
        queueArray.add(27, "PS GENView Analytics");
        queueArray.add(28, "PS GVBM");
        queueArray.add(29, "PS GVPP");
        queueArray.add(30, "PS GWC");
        queueArray.add(31, "PS hiG Gateways");
        queueArray.add(32, "PS IN");
        queueArray.add(33, "PS Kandy");
        queueArray.add(34, "PS Kandy Wrappers");
        queueArray.add(35, "PS LI / TOPS");
        queueArray.add(36, "PS Lines Services");
        queueArray.add(37, "PS MG15K G2 G6");
        queueArray.add(38, "PS MG9K");
        queueArray.add(39, "PS NSP");
        queueArray.add(40, "PS OAM IEMS");
        queueArray.add(41, "PS OAM SESM");
        queueArray.add(42, "PS OAM SPFS");
        queueArray.add(43, "PS Ribbon Protect");
        queueArray.add(44, "PS RSM");
        queueArray.add(45, "PS SBC");
        queueArray.add(46, "PS SeGW");
        queueArray.add(47, "PS Signaling");
        queueArray.add(48, "PS SIP Lines/SIP PBX");
        queueArray.add(49, "PS SPiDR CallP");
        queueArray.add(50, "PS SPiDR OAM");
        queueArray.add(51, "PS SPM MG4K");
        queueArray.add(52, "PS SST");
        queueArray.add(53, "PS Trunking");
        queueArray.add(54, "PS UT-SD");
        queueArray.add(55, "PS XLA");
        queueArray.add(56, "PS XPM V52");
        queueArray.add(57, "Tech-Ops ER Support");
        queueArray.add(58, "TS Asia");
        queueArray.add(59, "TS CALA");
        queueArray.add(60, "TS Converged Intelligent Messaging (CIM)");
        queueArray.add(61, "TS EDGE");
        queueArray.add(62, "TS EMEA");
        queueArray.add(63, "TS EMEA Marquee");
        queueArray.add(64, "TS EMEA PI");
        queueArray.add(65, "TS GTAC SERVICES");
        queueArray.add(66, "TS Japan Marquee");
        queueArray.add(67, "TS MEXICO");
        queueArray.add(68, "TS MNOC");
        queueArray.add(69, "TS NA");
        queueArray.add(70, "TS NA C15");
        queueArray.add(71, "TS NA DCO");
        queueArray.add(72, "TS NA Federal");
        queueArray.add(73, "TS NA G-Series");
        queueArray.add(74, "TS NA GTD5-5ESS");
        queueArray.add(75, "TS NA Marquee");
        queueArray.add(76, "TS NA Safari");
        queueArray.add(77, "TS NA Safari(GPS)");
        queueArray.add(78, "TS NA S-Series");
        queueArray.add(79, "TS NA Verizon Wireless");
        queueArray.add(80, "TS Non Technical");
        queueArray.add(81, "TS NSP");
        queueArray.add(82, "TS PSD");
        queueArray.add(83, "TS TAC-RESPONSE");
        queueArray.add(84, "TS TAQUA");
        queueArray.add(85, "TS UT-SD");
    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {

        readDefaultSettingFiles();
        setqueueArray();
        readTimeStamp();

    }
}