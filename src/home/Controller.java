package home;
import netscape.javascript.JSObject;

import com.jcraft.jsch.*;
import com.univocity.parsers.csv.CsvParser;
import com.univocity.parsers.csv.CsvParserSettings;
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
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.MenuItem;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.Image;
import javafx.scene.input.*;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.Pane;
import javafx.scene.text.Text;
import javafx.scene.web.WebEngine;
import javafx.scene.web.WebView;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Duration;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.awt.*;
import javafx.scene.control.TextArea;
import org.apache.poi.ss.usermodel.Font;

import java.io.*;
import java.net.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.logging.*;

public class Controller implements Initializable {

    //Top Frame Definitions
    @FXML
    private Label lblStatus;
    @FXML
    private FontAwesomeIconView btnUnlock;
    @FXML
    private FontAwesomeIconView btnToExcel;
    @FXML
    private FontAwesomeIconView btnBack;
    @FXML
    private FontAwesomeIconView btnInfo;

    //Left Frame Definitions
    @FXML
    private Button btnHome;
    @FXML
    private Button btnCases;
    @FXML
    private Button btnMyRegion;
    @FXML
    private Button btnMyNotes;
    @FXML
    private Button btnProducts;
    @FXML
    private Button btnCustomers;
    @FXML
    private Button btnProjects;
    @FXML
    private Button btnSettings;
    @FXML
    private Button btnProjection;
    @FXML
    private Button btnSkillSet;
    @FXML
    private Button btnLoadData;
    @FXML
    private Button btnLogin;
    @FXML
    private Label lblRefreshText;
    @FXML
    private Label lblDownload;
    @FXML
    private ProgressBar progressBar;

    //Overview Page:
    @FXML
    private AnchorPane apnHome;

    public final static Logger logger = Logger.getLogger("CMT_Main_Logger");
    FileHandler fh;

    //Account Page
    @FXML
    private AnchorPane apnAccounts;
    @FXML
    private Button btnAccE1Cases;
    @FXML
    private Button btnAccE2Cases;
    @FXML
    private Button btnAccOutFollow;
    @FXML
    private Button btnAccEscalated;
    @FXML
    private Button btnAccHotIssues;
    @FXML
    private Button btnAccBCCases;
    @FXML
    private Button btnAccInactive;
    @FXML
    private Button btnAccWOH;
    @FXML
    private Button btnAccBCupdated;
    @FXML
    private Button btnAccBCWac;
    @FXML
    private Button btnAccBCWIP;
    @FXML
    private Button btnAccBCINACT;
    @FXML
    private Button btnAccBCEngineering;
    @FXML
    private Button btnAccMJupdated;
    @FXML
    private Button btnAccMJWac;
    @FXML
    private Button btnAccMJWIP;
    @FXML
    private Button btnAccMJINACT;
    @FXML
    private Button btnAccMJEngineering;
    @FXML
    private Button btnAccMJDue;
    @FXML
    private Button btnAccBCMissed;
    @FXML
    private Button btnAccBCDue;
    @FXML
    private Button btnAccBCQueue;
    @FXML
    private Button btnAccMJQueue;
    @FXML
    private Button btnRegMJQueue;
    @FXML
    private Button btnAccMNQueue;
    @FXML
    private Button btnAccMNMissed;
    @FXML
    private Button btnAccMNDue;
    @FXML
    private Button btnAccMNINACT;
    @FXML
    private Button btnAccMNEngineering;
    @FXML
    private Button btnAccMNupdated;
    @FXML
    private Button btnAccMNWac;
    @FXML
    private Button btnAccMNWIP;
    @FXML
    private Button btnAccMJMissed;
    @FXML
    private Button btnaccGPSQueue;
    @FXML
    private Button btnAccRTSQueue;
    @FXML
    private Button btnAccUpdateToday;
    @FXML
    private Button btnAccUpdateMissed;
    @FXML
    private Button btnAccUpdateNull;
    @FXML
    private Pane pnAccSelect;
    @FXML
    private Pane pnWGSelect;
    @FXML
    private Pane pnWGNames;
    @FXML
    private TextField txAccounts;
    @FXML
    private TextArea txtWGNames;
    @FXML
    private TableView<AccountTableView> tableAcc;
    @FXML
    private TableView<WorkGroupTableView> tableWG;
    @FXML
    private TableView<WorkGroupTableView> tableWGSelected;
    @FXML
    private TableColumn<AccountTableView, String> accCol;
    @FXML
    private TableColumn<WorkGroupTableView, String> WGCol;
    @FXML
    private TableColumn<WorkGroupTableView, String> WGColSelected;
    @FXML
    private Button btnAccUpdate;
    @FXML
    private TableView<AccountTableView> tableAccSelected;
    @FXML
    private TableColumn<AccountTableView, String> accColSelected;
    @FXML
    private TextField txtAccSelect;
    @FXML
    private Button btnAccSelectClose;
    @FXML
    private Button btnAccSelectClear;
    @FXML
    private Button btnAccUpdateAdd;
    @FXML
    private Button btnAccUpdateRemove;
    @FXML
    private TextField txtWGSelect;
    @FXML
    private Button btnWGSelectClose;
    @FXML
    private Button btnWGSelectClear;
    @FXML
    private Button btnWGUpdateAdd;
    @FXML
    private Button btnWGUpdateRemove;
    @FXML
    private Button btnWGUpdate;

    // Region Wars
    @FXML
    private AnchorPane apnRegCases;
    @FXML
    private Button btnRegE1Cases;
    @FXML
    private Button btnRegE2Cases;
    @FXML
    private Button btnRegOutFollow;
    @FXML
    private Button btnRegEscalated;
    @FXML
    private Button btnRegHotIssues;
    @FXML
    private Button btnRegBCCases;
    @FXML
    private Button btnRegInactive;
    @FXML
    private Button btnRegWOH;
    @FXML
    private Button btnRegBCupdated;
    @FXML
    private Button btnRegBCWac;
    @FXML
    private Button btnRegBCWIP;
    @FXML
    private Button btnRegBCINACT;
    @FXML
    private Button btnRegBCEngineering;
    @FXML
    private Button btnRegMJupdated;
    @FXML
    private Button btnRegMJWac;
    @FXML
    private Button btnRegMJWIP;
    @FXML
    private Button btnRegMJINACT;
    @FXML
    private Button btnRegMJEngineering;
    @FXML
    private Button btnRegMJDue;
    @FXML
    private Button btnRegBCMissed;
    @FXML
    private Button btnRegBCDue;
    @FXML
    private Button btnRegMJMissed;
    @FXML
    private Button btnRegMNMissed;
    @FXML
    private Button btnRegUpdateToday;
    @FXML
    private Button btnRegUpdateMissed;
    @FXML
    private Button btnRegUpdateNull;
    @FXML
    private Button btnRegRTSQueue;
    @FXML
    private Button btnRegGPSQueue;
    @FXML
    private Button btnRegBCQueue;
    @FXML
    private Button btnRegMNupdated;
    @FXML
    private Button btnRegMNWac;
    @FXML
    private Button btnRegMNWIP;
    @FXML
    private Button btnRegMNINACT;
    @FXML
    private Button btnRegMNEngineering;
    @FXML
    private Button btnRegMNDue;
    @FXML
    private Button btnRegMNQueue;
    @FXML
    private ChoiceBox<String> regChoice;
    @FXML
    private Button btnForOverAll;
    @FXML
    private Button btnForMM;
    @FXML
    private Button btnForIMS;
    @FXML
    private AnchorPane apnWorkGroup;
    @FXML
    private WebView webWorkGroup;
    @FXML
    private AnchorPane apnForecastSel;
    @FXML
    private TextField forecastSelect;
    @FXML
    private ListView<String> lstForecast;
    @FXML
    private TextField forecastProductSelect;
    @FXML
    private RadioButton forecastAll;
    @FXML
    private Button btnForecastRun;
    @FXML
    private TextField txtUsersSave;
    @FXML
    private TextField txtProductsSave;
    @FXML
    private TextField txtQueuesSave;
    @FXML
    private TextField txtAccountSave;
    @FXML
    private Button btnUsersSaveAs;
    @FXML
    private Button btnProductsSaveAs;
    @FXML
    private Button btnQueuesSaveAs;
    @FXML
    private Button btnAccSaveAs;
    @FXML
    private Button btnUsersSaveClose;
    @FXML
    private Button btnUsersLoad;
    @FXML
    private Button btnProductsLoad;
    @FXML
    private Button btnQueuesLoad;
    @FXML
    private Button btnAccLoad;
    @FXML
    private Button btnUserProfDelete;
    @FXML
    private Button btnProductProfDelete;
    @FXML
    private Button btnQueueProfDelete;
    @FXML
    private Button btnAccountProfDelete;
    @FXML
    private Button btnProductsSave;
    @FXML
    private Button btnProductsSaveClose;
    @FXML
    private Button btnQueuesSave;
    @FXML
    private Button btnAccountSave;
    @FXML
    private Button btnUsersSave;
    @FXML
    private Button btnQueuesSaveClose;
    @FXML
    private Button btnAccountSaveClose;
    @FXML
    private Pane pnUsersSave;
    @FXML
    private Pane pnProductsSave;
    @FXML
    private Pane pnQueuesSave;
    @FXML
    private Pane pnAccountSave;
    @FXML
    private Pane pnUsersLoad;
    @FXML
    private Pane pnProductsLoad;
    @FXML
    private Pane pnQueuesLoad;
    @FXML
    private Pane pnAccountLoad;
    @FXML
    private AnchorPane spnNote;
    @FXML
    private AnchorPane apnProjection;
    @FXML
    private Pane pnCaseDetailsNote;
    @FXML
    private AnchorPane apnTableView;
    @FXML
    private AnchorPane apnManLogin;
    @FXML
    private AnchorPane apnNotes;
    @FXML
    private AnchorPane apnSettings;
    @FXML
    private AnchorPane apnSkills;
    @FXML
    private AnchorPane apnMyCases;
    @FXML
    private AnchorPane apnProduct;
    @FXML
    private AnchorPane apnProjects;
    @FXML
    private AnchorPane apnCustomers;
    @FXML
    private AnchorPane apnBrowser;
    @FXML
    private Pane browserLoginPane;
    @FXML
    private TextField txUsers;
    @FXML
    private TextField txProducts;
    @FXML
    public TextField customerText;
    @FXML
    private TextField txQueues;
    @FXML
    private TextField txWorkGroup;
    @FXML
    Text txtEscalated;
    @FXML
    private TextField txtCaseRegionNote;
    @FXML
    private TextField txtCaseAccountNote;
    @FXML
    private TextField txtCaseSubjectNote;
    @FXML
    private TextField txtCaseTypeNote;
    @FXML
    private TextField txtCaseStatusNote;
    @FXML
    private TextField txtCaseAgeNote;
    @FXML
    private TextField txtCaseOwnerNote;
    @FXML
    private TextField txtCaseCoOwnerNote;
    @FXML
    private TextField txtCaseCoQueueNote;
    @FXML
    private TextField txtCaseSeverityNote;
    @FXML
    private TextField txtCaseNumnberNote;
    @FXML
    private CheckBox checkBoxHotIssueNote;
    @FXML
    private CheckBox checkBoxEscalatedNote;
    @FXML
    private CheckBox checkRegProduct;
    @FXML
    private TextField txtCaseProductNote;
    @FXML
    private Button btnProjectRight;
    @FXML
    private Button btnProjectLeft;
    @FXML
    private Button btnDelNote;
    @FXML
    private Button btnAddNewNote;
    @FXML
    private Button btnViewNote;
    @FXML
    private Button btnViewComment;
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
    private Button btnMyBCQueue;
    @FXML
    private Button btnMyMJQueue;
    @FXML
    private Button btnMyMNupdated;
    @FXML
    private Button btnMyMNWac;
    @FXML
    private Button btnMyMNWIP;
    @FXML
    private Button btnMyMNINACT;
    @FXML
    private Button btnMyMNEngineering;
    @FXML
    private Button btnMyMNDue;
    @FXML
    private Button btnMyMNQueue;
    @FXML
    private Button btnMyMNMissed;
    @FXML
    private Button btnBCMissed;
    @FXML
    private Button btnBCQueue;
    @FXML
    private Button btnMyBCMissed;
    @FXML
    private Button btnMJDue;
    @FXML
    private Button btnMyMJDue;
    @FXML
    private Button btnMJMissed;
    @FXML
    private Button btnMJQueue;
    @FXML
    private Button btnMNWIP;
    @FXML
    private Button btnMNWac;
    @FXML
    private Button btnMNupdated;
    @FXML
    private Button btnMNEngineering;
    @FXML
    private Button btnMNINACT;
    @FXML
    private Button btnMNDue;
    @FXML
    private Button btnMNMissed;
    @FXML
    private Button btnMNQueue;
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
    private Button btnMyCoOwnQueue;
    @FXML
    private Button btnMyCoQueueAssigned;
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
    private Button btnBCQueueProd;
    @FXML
    private Button btnMJQueueProd;
    @FXML
    private Button btnMNupdatedProd;
    @FXML
    private Button btnMNWacProd;
    @FXML
    private Button btnMNWIPProd;
    @FXML
    private Button btnMNINACTProd;
    @FXML
    private Button btnMNEngineeringProd;
    @FXML
    private Button btnMNDueProd;
    @FXML
    private Button btnMNQueueProd;
    @FXML
    private Button btnMNMissedProd;
    @FXML
    private Button btnAccountClear;
    @FXML
    private Button btnAccClear;
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
    private Button btnAmericas;
    @FXML
    private Button btnEmea;
    @FXML
    private Button btnApac;
    @FXML
    private Button btnJapan;
    @FXML
    private Button btnGatingNow;
    @FXML
    private Button btnGatingDate;
    @FXML
    private Button btnGatingPrevious;
    @FXML
    private Button btnProjectsAll;
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
    private TableColumn<CaseTableView, String> CoOwnerCol;
    @FXML
    private TableColumn<CaseTableView, String> CoOwnerQueueCol;
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
    private TableColumn<CaseTableView, String> CoOwnerColCust;
    @FXML
    private TableColumn<CaseTableView, String> CoOwnerQueuColCust;
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
    private Button btnFilterAccountUpdate;
    @FXML
    private Button btnFilterAccountUpdateAdd;
    @FXML
    private Button btnFilterAccountUpdateRemove;
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
    private Button btnUsersUpdateAdd;
    @FXML
    private Button btnUsersUpdateRemove;
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
    private Button btnWGClear;
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
    private Button btnProductUpdateAdd;
    @FXML
    private Button btnProductUpdateRemove;
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
    private TableView<ProjectTableView> tableProjects;
    @FXML
    private TableColumn<ProjectTableView, String> prjNumCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjNoCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjSevCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjStatCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjProdCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjAccCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjSubCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjModCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjHotRCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjGateDateCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjRegionCol;
    @FXML
    private TableColumn<ProjectTableView, String> prjSiteStatusCol;
    @FXML
    private Button btnQueueUpdate;
    @FXML
    private Button btnQueueUpdateAdd;
    @FXML
    private Button btnQueueUpdateRemove;
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
    @FXML
    private Button btnUserProfLoad;
    @FXML
    private Button btnUsersLoadClose;
    @FXML
    private Button btnProdProfLoad;
    @FXML
    private Button btnProductsLoadClose;
    @FXML
    private Button btnQueueProfLoad;
    @FXML
    private Button btnQueueLoadClose;
    @FXML
    private Button btnAccountProfLoad;
    @FXML
    private Button btnAccountLoadClose;
    @FXML
    private ListView caseNoteList;
    @FXML
    private ListView userProfileList;
    @FXML
    private ListView productProfileList;
    @FXML
    private ListView queueProfileList;
    @FXML
    private ListView accountProfileList;
    @FXML
    private TextArea txtShowCaseNotes;
    @FXML
    private Button btnPrjMyNotes;
    @FXML
    private Button btnPrjNewNote;
    @FXML
    private Button btnPrjDelNote;
    @FXML
    private Pane pnPrjNotes;
    @FXML
    private ListView lstPrjNotes;
    @FXML
    private TextArea txtPrjNoteView;
    @FXML
    private WebView webviewTest;
    @FXML
    private WebView wvWG;
    @FXML
    private WebView projectWeb;
    @FXML
    private TextField txtpass;
    @FXML
    private Button btnManLogin;
    @FXML
    private Button btnManClose;
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
    private Button btnSkillsExport;

    ArrayList<String> readUserList;
    ArrayList<String> readOverAllUsers;
    ArrayList<String> safeUserList;
    int userSkillRef;
    int skillRef;
    String selectedLevel;
    String selected;
    String selectedSkill;
    String selectedSkillLevel;
    String selectedRegion;
    ArrayList<String> expertLevel;
    ArrayList<String> intLevel;
    ArrayList<String> basicLevel;
    ArrayList<String> noLevel;
    ArrayList<String> skillsAll;
    ArrayList<String> skillsExpert;
    ArrayList<String> skillsInterm;
    ArrayList<String> skillsBegin;
    ArrayList<String> skillsNone;

    Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
    int screenHeight = screenSize.height;

    ObservableList<String> levels = FXCollections.observableArrayList();

    static String rootFolder;
    static String settingsFolder;
    static String dataFolder;
    static String selectionFolder;
    static String skillSetFolder;
    static String logFolder;
    static String noteFolder;
    static String detailsFolder;
    static File repo = new File(System.getProperty("user.home") + "\\Documents\\CMT");

    Timeline timeData = new Timeline();
    CMTFolder folderArrange = new CMTFolder();
    DownData downData = new DownData();

    WebView browserLogin = new WebView();
    ArrayList<String> settingsUsers = new ArrayList<>();
    ArrayList<String> settingsQueue = new ArrayList<>();
    ArrayList<String> settingsProducts = new ArrayList<>();
    ArrayList<String> settingsAccounts = new ArrayList<>();
    ArrayList<String> settingsWorkGroups = new ArrayList<>();
    ArrayList<String> filteredAccounts = new ArrayList<String>();
    ArrayList<String> filteredWGs = new ArrayList<String>();
    ArrayList<String> WgFiltered = new ArrayList<String>();
    ArrayList<String> setWG = new ArrayList<String>();
    ArrayList<String> setWgNames = new ArrayList<String>();
    ArrayList<String> setUsers = new ArrayList<String>();
    ArrayList<String> usersFiltered = new ArrayList<String>();
    ArrayList<String> productsFiltered = new ArrayList<String>();
    ArrayList<String> accountsFiltered = new ArrayList<String>();
    ArrayList<String> queuesFiltered = new ArrayList<String>();
    ArrayList<String> queueArray = new ArrayList<>();
    ArrayList<String> selectedCase = new ArrayList<>();

    ContextMenu menu = new ContextMenu();
    MenuItem openCaseSFDC = new MenuItem("Search This Case in SalesForce...");
    MenuItem casePersonalNote = new MenuItem("Add Memo For This Case...");
    MenuItem openCaseComments = new MenuItem("Read Work Notes for this case...");
    MenuItem openCaseDetails = new MenuItem("View Details...");

    //Case Ref Cells
    int caseWGRef = 0;
    int caseWGNameRef = 0;
    int caseVPref = 0;
    int caseAccountRef = 0;
    int caseNumCellRef = 0;
    int caseSupTypeRefCell = 0;
    int caseStatRefCell = 0;
    int caseSevRefCell = 0;
    int caseRespRefCell = 0;
    int caseOwnerRefCell = 0;
    int caseCoOwnerRefCell = 0;
    int caseEscalatedRefCell = 0;
    int caseSupHotListRRef = 0;
    int caseRegionRef;
    int rbnDaysRegRefCell = 0;
    int caseHotListRefCell = 0;
    int caseOutFolRefCell = 0;
    int caseAgeRefCell = 0;
    int rbbnDaysRefCell = 0;
    int mycaseNumCellRef = 0;
    int mycaseSupTypeRefCell = 0;
    int mycaseStatRefCell = 0;
    int mycaseSevRefCell = 0;
    int mycaseRegRefCell = 0;
    int mycaseRespRefCell = 0;
    int mycaseOwnerRefCell = 0;
    int mycaseEscalatedRefCell = 0;
    int mycaseHotListRefCell = 0;
    int mycaseOutFolRefCell = 0;
    int mycaseAgeRefCell = 0;
    int mycaseUpdateCell = 0;
    int myCoOwnCaseRefCell = 0;
    int myCoOwnQueueRefCell = 0;
    int myRbbnDaysRefCell = 0;

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
    int rbnDaysRefCell=0;

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
    int bcQueue = 0;
    int mjQueue = 0;
    int mnQueue = 0;
    int dueMJday = 0;
    int dueMNday = 0;
    int misMJdue = 0;
    int misMNdue = 0;
    int custActMJ = 0;
    int custActMN = 0;
    int custRpdMJ = 0;
    int custRpdMN = 0;
    int MJds = 0;
    int MNpc = 0;
    int MNds = 0;
    int MJpc = 0;

    int MJwip = 0;
    int MNwip = 0;

    //Accounts Page Variables

    int accHotList;
    int accOutFollow;
    int accEscCases;
    int accBCCases;
    int accInactiveCases;
    int accBCDueCases;
    int accBCMissedCases;
    int accBCDSCases;
    int accBCInactiveCases;
    int accBCWIP;
    int accMJDueCases;
    int accMJMissedCases;
    int accMNMissedCases;
    int accRTSQueue;
    int accGPSQueue;
    int accMJUpdated;
    int accMJDSCases;
    int accMJWIP;
    int accQueuedCases;
    int accCoOwnerQueueCases;
    int accCoOwnerQueueCasesAssigned;
    int accE1Case;
    int accE2Cases;
    int accBCupdated;
    int accBCWac;
    int accMJWAC;
    int accMJInactiveCases;
    int accWOHCases;
    int accUpdateToday;
    int accUpdateMissed;
    int accUpdateNull;
    int accCoOwnCase;
    int accCoOwnQueue;
    int accBCQueue;
    int accMJQueue;
    int accMNQueue;
    int accMNWIP;
    int accMNWAC;
    int accMNUpdated;
    int accMNEng;
    int accMNInact;
    int accMNDue;

    //Region Page variables

    int regHotList;
    int regOutFollow;
    int regEscCases;
    int regBCCases;
    int regInactiveCases;
    int regBCDueCases;
    int regBCMissedCases;
    int regBCDSCases;
    int regBCInactiveCases;
    int regBCWIP;
    int regMJDueCases;
    int regMJMissedCases;
    int regMNMissedCases;
    int regRTSQueue;
    int regGPSQueue;
    int regMJUpdated;
    int regMJDSCases;
    int regMJWIP;
    int regQueuedCases;
    int regCoOwnerQueueCases;
    int regCoOwnerQueueCasesAssigned;
    int regE1Case;
    int regE2Cases;
    int regBCupdated;
    int regBCWac;
    int regMJWAC;
    int regMJInactiveCases;
    int regWOHCases;
    int regUpdateToday;
    int regUpdateMissed;
    int regUpdateNull;
    int regCoOwnCase;
    int regCoOwnQueue;
    int regBCQueue;
    int regMJQueue;
    int regMNQueue;
    int regMNWIP;
    int regMNWAC;
    int regMNUpdated;
    int regMNEng;
    int regMNInact;
    int regMNDue;
    int regMNMissed;

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
    int myCoOwnerQueueCases = 0;
    int myCoOwnerQueueCasesAssigned = 0;
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
    int myCoOwnCase = 0;
    int myCoOwnQueue = 0;
    int myBCInQueue = 0;
    int myMJInQueue = 0;
    int myMNInQueue = 0;
    int myMNUpdated = 0;
    int myMNWAC = 0;
    int myMNWIP = 0;
    int myMNEng = 0;
    int myMNDue = 0;
    int myMNMissed = 0;
    int myMNINAct = 0;

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
    int prodBCQueue = 0;
    int prodMJQueue = 0;
    int prodMNQueue = 0;
    int prodMNWIP = 0;
    int prodMNWAC = 0;
    int prodMNInact = 0;
    int prodMNUpdated = 0;
    int prodMNEng = 0;
    int prodMNDue = 0;
    int prodMNMissed = 0;

    //Project Page Variables
    int prjAmericas = 0;
    int prjEmea = 0;
    int prjApac = 0;
    int prjJapan = 0;
    int prjGatingNow = 0;
    int prjGatingDate = 0;
    int prjGatingPrev = 0;
    int prjAllCases = 0;
    boolean reg_product = false;

    //Alert Strings
    String strAlert = "NO RECORD FOUND...";
    String strNoNote = "There is No Personal Memo Saved..." + "\n" + "\n" + "PLEASE CREATE PERSONAL MEMO FIRST!";
    String strSave = "PLEASE PROMPT A PROFILE NAME";
    String strLoadProf = "THERE IS NO USER PROFILE SAVED!";


    @FXML
    void handleAccClicks(ActionEvent event) {

        //Create the Account Based Tables on related tile click
        if (event.getSource() == btnAccE1Cases){
            if (accE1Case != 0){
                lblStatus.setText("CRITICAL CASES (" + txAccounts.getText()+")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter1 = "Critical";
                initTableView(tableCases);
                accOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccE2Cases){
            if (accE2Cases != 0){
                lblStatus.setText("E2 CASES (" + txAccounts.getText()+")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter1 = "E2";
                initTableView(tableCases);
                accOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccOutFollow){
            if (accOutFollow != 0){
                lblStatus.setText("OUTAGE FOLLOW-UP CASES (" + txAccounts.getText()+")");
                tableCases.getItems().clear();
                String columnSelect = "Outage Follow-Up";
                String filter1 = "1";
                initTableView(tableCases);
                accOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccEscalated){
            if (accEscCases != 0){
                lblStatus.setText("ESCALATED CASES (" + txAccounts.getText()+")");
                tableCases.getItems().clear();
                String columnSelect = "Escalated By";
                String filter1 = "NotSet";
                initTableView(tableCases);
                accOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, false);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccBCCases){
            if (accBCCases != 0){
                lblStatus.setText("BC CASES (" + txAccounts.getText()+")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter1 = "Business Critical";
                initTableView(tableCases);
                accOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccHotIssues){
            if (accHotList != 0){
                lblStatus.setText("HOT ISSUES (" + txAccounts.getText()+")");
                tableCases.getItems().clear();
                String columnSelect = "Support Hotlist Level";
                String filter1 = "NotSet";
                initTableView(tableCases);
                accOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, false);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccWOH){
            if (accWOHCases != 0){
                lblStatus.setText("ACTIVE CASES (" + txAccounts.getText()+")");
                tableCases.getItems().clear();
                initTableView(tableCases);
                accOverviewWOHView(true);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccInactive){
            if (accInactiveCases != 0){
                lblStatus.setText("INACTIVE CASES (" + txAccounts.getText()+")");
                tableCases.getItems().clear();
                initTableView(tableCases);
                accOverviewWOHView(false);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccBCWIP){
            if (accBCWIP != 0){
                lblStatus.setText("BC WORK IN PROGRESS CASES (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                accWIPCaseTableView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccBCWac) {
            if (accBCWac != 0) {
                lblStatus.setText("BC CASES PENDING CUSTOMER ACTION (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Business Critical";
                String filter2 = "Customer action";
                initTableView(tableCases);
                accTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccBCupdated) {
            if (accBCupdated != 0) {
                lblStatus.setText("BC CASES PENDING OWNER ACTION ("+ txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Business Critical";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                accTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccBCEngineering) {
            if (accBCDSCases != 0) {
                lblStatus.setText("BC CASES WITH ENGINEERING (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columSelect = "Severity";
                String filter1 = "Business Critical";
                initTableView(tableCases);
                accEngineeringTableView(columSelect, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccBCINACT) {
            if (accBCInactiveCases != 0) {
                lblStatus.setText("INACTIVE BC CASES (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String filter1 = "Business Critical";
                initTableView(tableCases);
                accInactiveTable(columnSelect1, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMJWIP) {
            if (accMJWIP != 0) {
                lblStatus.setText("MAJOR WORK IN PROGRESS CASES (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                accWIPCaseTableView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMNWIP) {
            if (accMNWIP != 0) {
                lblStatus.setText("MINOR WORK IN PROGRESS CASES (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                accWIPCaseTableView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMJWac) {
            if (accMJWAC != 0) {
                lblStatus.setText("MAJOR CASES WITH CUSTOMER (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Major";
                String filter2 = "Customer action";
                initTableView(tableCases);
                accTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMNWac) {
            if (accMNWAC != 0) {
                lblStatus.setText("MINOR CASES WITH CUSTOMER (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Minor";
                String filter2 = "Customer action";
                initTableView(tableCases);
                accTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMJupdated) {
            if (accMJUpdated != 0) {
                lblStatus.setText("MAJOR CASES PENDING OWNER ACTION (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Major";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                accTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMNupdated) {
            if (accMNUpdated != 0) {
                lblStatus.setText("MINOR CASES PENDING OWNER ACTION (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Minor";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                accTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMJEngineering) {
            if (accMJDSCases != 0) {
                lblStatus.setText("MAJOR CASES WITH ENGINEERING (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columSelect = "Severity";
                String filter1 = "Major";
                initTableView(tableCases);
                accEngineeringTableView(columSelect, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMNEngineering) {
            if (accMNEng != 0) {
                lblStatus.setText("MINOR CASES WITH ENGINEERING (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columSelect = "Severity";
                String filter1 = "Minor";
                initTableView(tableCases);
                accEngineeringTableView(columSelect, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMJINACT) {
            if (accMJInactiveCases != 0) {
                lblStatus.setText("INACTIVE MAJOR CASES (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String filter1 = "Major";
                initTableView(tableCases);
                accInactiveTable(columnSelect1, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMNINACT) {
            if (accMNInact != 0) {
                lblStatus.setText("INACTIVE MINOR CASES (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String filter1 = "Minor";
                initTableView(tableCases);
                accInactiveTable(columnSelect1, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccBCDue) {
            if (accBCDueCases != 0) {
                lblStatus.setText("BC DUE CASES (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                accDueFilterView(columnSelect, filter, tableCases, apnTableView, 5, true);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccBCMissed) {
            if (accBCMissedCases != 0) {
                lblStatus.setText("BC CASES MISSED DUE (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                accDueFilterView(columnSelect, filter, tableCases, apnTableView, 5, false);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMJDue) {
            if (accMJDueCases != 0) {
                lblStatus.setText("MAJOR DUE CASES (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                accDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, true);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMNDue) {
            if (accMNDue != 0) {
                lblStatus.setText("MINOR DUE CASES (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                accDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, true);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMJMissed) {
            if (accMJMissedCases != 0) {
                lblStatus.setText("MAJOR CASES MISSED DUE (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                accDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, false);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMNMissed) {
            if (accMNMissedCases != 0) {
                lblStatus.setText("MINOR CASES MISSED DUE (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                accDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, false);
                tableCases.scrollToColumnIndex(0);
            }
        }

        if (event.getSource() == btnAccRTSQueue) {
            if (accRTSQueue != 0) {
                lblStatus.setText("CASES IN RTS QUEUE FOR: " + txAccounts.getText());
                tableCases.getItems().clear();
                String columnselect = "Case Owner";
                String filter = "TS";
                initTableView(tableCases);
                accQueueView(columnselect, filter, tableCases, apnTableView, "Tech-Ops", "All");
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnaccGPSQueue) {
            if (accGPSQueue != 0) {
                lblStatus.setText("CASES IN GPS QUEUE FOR: " + txAccounts.getText());
                tableCases.getItems().clear();
                String e2TableSelect = "Case Owner";
                String e2TableSelect2 = "PS";
                initTableView(tableCases);
                accQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView, "GPS", "All");
                tableCases.scrollToColumnIndex(0);
            }
        }

        if (event.getSource() == btnAccBCQueue) {
            if (accBCQueue != 0) {
                lblStatus.setText("BC CASES IN GPS QUEUE For (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String e2TableSelect = "Case Owner";
                String e2TableSelect2 = "PS";
                initTableView(tableCases);
                accQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView, "GPS", "BC");
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMJQueue) {
            if (accMJQueue != 0) {
                lblStatus.setText("MAJOR CASES IN GPS QUEUE For (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String e2TableSelect = "Case Owner";
                String e2TableSelect2 = "PS";
                initTableView(tableCases);
                accQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView, "GPS", "Major");
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccMNQueue) {
            if (accMNQueue != 0) {
                lblStatus.setText("MINOR CASES IN GPS QUEUE For (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String e2TableSelect = "Case Owner";
                String e2TableSelect2 = "PS";
                initTableView(tableCases);
                accQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView, "GPS", "Minor");
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccUpdateToday) {
            if (accUpdateToday != 0) {
                lblStatus.setText("NEXT CASE UPDATE TODAY (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                accCaseUpdateTableView(columnSelect, tableCases, apnTableView, true, true);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccUpdateMissed) {
            if (accUpdateMissed != 0) {
                lblStatus.setText("NEXT CASE UPDATE MISSED ("+ txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                accCaseUpdateTableView(columnSelect, tableCases, apnTableView, false, true);
                tableCases.scrollToColumnIndex(0);
            }
        }
        if (event.getSource() == btnAccUpdateNull) {
            if (accUpdateNull != 0) {
                lblStatus.setText("NEXT CASE UPDATE FIELD NOT SET CASES (" + txAccounts.getText() + ")");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                accCaseUpdateTableView(columnSelect, tableCases, apnTableView, false, false);
                tableCases.scrollToColumnIndex(0);
            }
        }
    }

    @FXML
    void handleRegClicks(ActionEvent event) {

        if (event.getSource() == btnRegE1Cases){

            if (regE1Case != 0) {
                lblStatus.setText("CRITICAL (OUTAGE) CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter1 = "Critical";
                initTableView(tableCases);
                regionOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (regE1Case == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnRegE2Cases) {
            if (regE2Cases != 0) {
                lblStatus.setText("E2 CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter1 = "E2";
                initTableView(tableCases);
                regionOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (regE2Cases == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnRegOutFollow) {
            if (regOutFollow != 0) {
                lblStatus.setText("OUTAGE FOLLOW-UP CASES FOR "+ selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Outage Follow-Up";
                String filter1 = "1";
                initTableView(tableCases);
                regionOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (regOutFollow == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegEscalated) {
            if (regEscCases != 0) {
                lblStatus.setText("ESCALATED CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Escalated By";
                String filter1 = "NotSet";
                initTableView(tableCases);
                regionOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (regEscCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegBCCases) {
            if (regBCCases != 0) {
                lblStatus.setText("BUSINESS CRITICAL CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter1 = "Business Critical";
                initTableView(tableCases);
                regionOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (regBCCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegHotIssues) {
            if (regHotList != 0) {
                lblStatus.setText("HOT ISSUES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Support Hotlist Level";
                String filter1 = "NotSet";
                initTableView(tableCases);
                regionOneFilterTableView(columnSelect, filter1, tableCases, apnTableView, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (regHotList == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegWOH) {
            if (regWOHCases != 0) {
                lblStatus.setText("ACTIVE WORK ON HAND CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                initTableView(tableCases);
                regOverviewWOHView(true);
                tableCases.scrollToColumnIndex(0);
            }
            if (regWOHCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegInactive) {
            if (regInactiveCases != 0) {
                lblStatus.setText("INACTIVE(PC/FA) CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                initTableView(tableCases);
                regOverviewWOHView(false);
                tableCases.scrollToColumnIndex(0);
            }
            if (regInactiveCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegBCWIP) {
            if (regBCWIP != 0) {
                lblStatus.setText("BUSINESS CRITICAL WORK IN PROGRESS CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                regionWIPCaseTableView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regBCWIP == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegBCWac) {
            if (regBCWac != 0) {
                lblStatus.setText("BC CASES PENDING CUSTOMER ACTION FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Business Critical";
                String filter2 = "Customer action";
                initTableView(tableCases);
                regTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regBCWac == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegBCupdated) {
            if (regBCupdated != 0) {
                lblStatus.setText("BC CASES PENDING OWNER ACTION FOR "+ selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Business Critical";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                regTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regBCupdated == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegBCEngineering) {
            if (regBCDSCases != 0) {
                lblStatus.setText("BC CASES WITH ENGINEERING FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columSelect = "Severity";
                String filter1 = "Business Critical";
                initTableView(tableCases);
                regEngineeringTableView(columSelect, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regBCDSCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegBCINACT) {
            if (regBCInactiveCases != 0) {
                lblStatus.setText("INACTIVE BC CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String filter1 = "Business Critical";
                initTableView(tableCases);
                regInactiveTable(columnSelect1, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regBCInactiveCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMJWIP) {
            if (regMJWIP != 0) {
                lblStatus.setText("MAJOR WORK IN PROGRESS CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                regionWIPCaseTableView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMJWIP == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnRegMNWIP) {
            if (regMNWIP != 0) {
                lblStatus.setText("MINOR WORK IN PROGRESS CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                regionWIPCaseTableView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMNWIP == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMJWac) {
            if (regMJWAC != 0) {
                lblStatus.setText("MAJOR CASES WITH CUSTOMER FOR" + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Major";
                String filter2 = "Customer action";
                initTableView(tableCases);
                regTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMJWAC == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMNWac) {
            if (regMNWAC != 0) {
                lblStatus.setText("MINOR CASES WITH CUSTOMER FOR" + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Minor";
                String filter2 = "Customer action";
                initTableView(tableCases);
                regTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMNWAC == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMJupdated) {
            if (regMJUpdated != 0) {
                lblStatus.setText("MAJOR CASES PENDING OWNER ACTION FOR" + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Major";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                regTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMJUpdated == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnRegMNupdated) {
            if (regMNUpdated != 0) {
                lblStatus.setText("MINOR CASES PENDING OWNER ACTION FOR" + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Minor";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                regTwoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMNWAC == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMJEngineering) {
            if (regMJDSCases != 0) {
                lblStatus.setText("MAJOR CASES WITH ENGINEERING FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columSelect = "Severity";
                String filter1 = "Major";
                initTableView(tableCases);
                regEngineeringTableView(columSelect, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMJDSCases == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnRegMNEngineering) {
            if (regMNEng != 0) {
                lblStatus.setText("MINOR CASES WITH ENGINEERING FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columSelect = "Severity";
                String filter1 = "Minor";
                initTableView(tableCases);
                regEngineeringTableView(columSelect, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMNEng == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMJINACT) {
            if (regMJInactiveCases != 0) {
                lblStatus.setText("INACTIVE MAJOR CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String filter1 = "Major";
                initTableView(tableCases);
                regInactiveTable(columnSelect1, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMJInactiveCases == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnRegMNINACT) {
            if (regMNInact != 0) {
                lblStatus.setText("INACTIVE MINOR CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String filter1 = "Minor";
                initTableView(tableCases);
                regInactiveTable(columnSelect1, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMNInact == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegBCDue) {
            if (regBCDueCases != 0) {
                lblStatus.setText("BC DUE CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                regDueFilterView(columnSelect, filter, tableCases, apnTableView, 5, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (regBCDueCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegBCMissed) {
            if (regBCMissedCases != 0) {
                lblStatus.setText("BC CASES MISSED DUE FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                regDueFilterView(columnSelect, filter, tableCases, apnTableView, 5, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (regBCMissedCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMJDue) {
            if (regMJDueCases != 0) {
                lblStatus.setText("MAJOR DUE CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                regDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMJDueCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMNDue) {
            if (regMNDue != 0) {
                lblStatus.setText("MAJOR DUE CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                regDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMNDue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMJMissed) {
            if (regMJMissedCases != 0) {
                lblStatus.setText("MAJOR CASES MISSED DUE FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                regDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMJMissedCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMNMissed) {
            if (regMNMissed != 0) {
                lblStatus.setText("MINOR CASES MISSED DUE FOR" + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                regDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMNMissed == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegRTSQueue) {
            if (regRTSQueue != 0) {
                lblStatus.setText("CASES IN RTS QUEUE FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnselect = "Case Owner";
                String filter = "TS";
                initTableView(tableCases);
                String severity = "All";
                regQueueView(columnselect, filter, tableCases, apnTableView, "Tech-Ops", severity);
                tableCases.scrollToColumnIndex(0);
            }
            if (regRTSQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegGPSQueue) {
            if (regGPSQueue != 0) {
                lblStatus.setText("CASES IN GPS QUEUE FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String e2TableSelect = "Case Owner";
                String e2TableSelect2 = "PS";
                initTableView(tableCases);
                String severity = "All";
                regQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView, "GPS", severity);
                tableCases.scrollToColumnIndex(0);
            }
            if (regGPSQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegBCQueue) {
            if (regBCQueue != 0) {
                lblStatus.setText("BC CASES IN GPS QUEUE FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String e2TableSelect = "Case Owner";
                String e2TableSelect2 = "Business Critical";
                initTableView(tableCases);
                String severity = "Business Critical";
                regQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView, "BC In Queue", severity);
                tableCases.scrollToColumnIndex(0);
            }
            if (regBCQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMJQueue) {
            if (regMJQueue != 0) {
                lblStatus.setText("MJ CASES IN GPS QUEUE FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String e2TableSelect = "Case Owner";
                String e2TableSelect2 = "Major";
                initTableView(tableCases);
                String severity = "Major";
                regQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView, "Major In Queue", severity);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMJQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegMNQueue) {
            if (regMNQueue != 0) {
                lblStatus.setText("MN CASES IN GPS QUEUE FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String e2TableSelect = "Case Owner";
                String e2TableSelect2 = "Minor";
                initTableView(tableCases);
                String severity = "Minor";
                regQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView, "Minor In Queue", severity);
                tableCases.scrollToColumnIndex(0);
            }
            if (regMNQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegUpdateToday) {
            if (regUpdateToday != 0) {
                lblStatus.setText("NEXT CASE UPDATE TODAY FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                regCaseUpdateTableView(columnSelect, tableCases, apnTableView, true, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (regUpdateToday == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegUpdateMissed) {
            if (regUpdateMissed != 0) {
                lblStatus.setText("NEXT CASE UPDATE MISSED FOR "+ selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                regCaseUpdateTableView(columnSelect, tableCases, apnTableView, false, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (regUpdateMissed == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnRegUpdateNull) {
            if (regUpdateNull != 0) {
                lblStatus.setText("NEXT CASE UPDATE FIELD NOT SET CASES FOR " + selectedRegion + " REGION");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                regCaseUpdateTableView(columnSelect, tableCases, apnTableView, false, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (regUpdateNull == 0) {
                alertUser(strAlert);
            }
        }
    }
    private void accQueueView(String columnSelect, String filter, TableView tableView, AnchorPane anchorpane, String overText, String sev) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;

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
                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Severity")){
                    caseSevRefCell = i;
                }
            }

            ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
            int accfiltnum = setAcc.size();

            for (int i = 0; i <accfiltnum ; i++) {

                for (int k = 1; k < lastRow + 1; k++) {

                    cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                    String cellValToCompare = cellVal.getStringCellValue();
                    cellVal2 = filtersheet.getRow(k).getCell(caseAccountRef);
                    String acc = cellVal2.getStringCellValue();
                    cellVal3 = filtersheet.getRow(k).getCell(rbnDaysRefCell);
                    String rbDays = cellVal3.getStringCellValue();
                    int ribDays = Integer.parseInt(rbDays);
                    cellVal4 = filtersheet.getRow(k).getCell(caseSevRefCell);
                    String severity = cellVal4.getStringCellValue();

                    if (acc.equals(setAcc.get(i))) {

                        switch (sev) {
                            case ("All"):
                                if (cellValToCompare.equals(filter) || cellValToCompare.startsWith(filter) || cellValToCompare.startsWith(overText)) {

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
                                            array.get(3), array.get(4), array.get(5), ribDays, age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableView.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableView.getItems().size() >= caseCount + 1) {
                                        tableView.getItems().removeAll(observableList);
                                    }
                                }
                                break;
                            case ("BC"):
                                if (severity.equals("Business Critical") && (cellValToCompare.startsWith("PS") || cellValToCompare.startsWith("TS") || cellValToCompare.startsWith("Tech-Ops"))) {

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
                                            array.get(3), array.get(4), array.get(5), ribDays, age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableView.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableView.getItems().size() >= caseCount + 1) {
                                        tableView.getItems().removeAll(observableList);
                                    }
                                }
                                break;
                            case ("Major"):
                                if (severity.equals("Major") && (cellValToCompare.startsWith("PS") || cellValToCompare.startsWith("TS") || cellValToCompare.startsWith("Tech-Ops"))) {

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
                                            array.get(3), array.get(4), array.get(5), ribDays, age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableView.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableView.getItems().size() >= caseCount + 1) {
                                        tableView.getItems().removeAll(observableList);
                                    }
                                }
                                break;
                            case("Minor"):

                                if (severity.equals("Minor") && (cellValToCompare.startsWith("PS") || cellValToCompare.startsWith("TS") || cellValToCompare.startsWith("Tech-Ops"))) {

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
                                            array.get(3), array.get(4), array.get(5), ribDays, age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableView.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableView.getItems().size() >= caseCount + 1) {
                                        tableView.getItems().removeAll(observableList);
                                    }
                                }
                                break;
                        }
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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnAccounts.toFront();
                    lblStatus.setText("ACCOUNT(S) VIEW ( " + txAccounts.getText() + ")");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }
    private void regQueueView(String columnSelect, String filter, TableView tableView, AnchorPane anchorpane, String overText, String severity) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;

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
                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Severity")){
                    caseSevRefCell = i;
                }

            }
            for (int k = 1; k < lastRow + 1; k++) {

                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();
                cellVal2 = filtersheet.getRow(k).getCell(caseRegionRef);
                String region = cellVal2.getStringCellValue();
                cellVal3 = filtersheet.getRow(k).getCell(rbnDaysRefCell);
                String rbDays = cellVal3.getStringCellValue();
                int ribDays = Integer.parseInt(rbDays);
                cellVal4 = filtersheet.getRow(k).getCell(caseSevRefCell);
                String sever = cellVal4.getStringCellValue();

                if (region.equals(selectedRegion) && (!overText.equals(""))) {

                    if (severity.equals("All")) {

                        if (cellValToCompare.equals(filter) || cellValToCompare.startsWith(filter) || cellValToCompare.startsWith(overText)) {

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
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableView.getItems().addAll(observableList);
                            caseCount++;
                            if (tableView.getItems().size() >= caseCount + 1) {
                                tableView.getItems().removeAll(observableList);
                            }
                        }
                    }
                    if (severity.equals("Business Critical")){

                        if(sever.equals("Business Critical") && (cellValToCompare.startsWith("PS") || cellValToCompare.startsWith("TS") || cellValToCompare.startsWith("Tech-Ops"))){

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
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableView.getItems().addAll(observableList);
                            caseCount++;
                            if (tableView.getItems().size() >= caseCount + 1) {
                                tableView.getItems().removeAll(observableList);
                            }
                        }

                    }
                    if (severity.equals("Major")){
                        if((sever.equals("Major") && (cellValToCompare.startsWith("PS") || cellValToCompare.startsWith("TS") || cellValToCompare.startsWith("Tech-Ops")))){

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
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableView.getItems().addAll(observableList);
                            caseCount++;
                            if (tableView.getItems().size() >= caseCount + 1) {
                                tableView.getItems().removeAll(observableList);
                            }
                        }
                    }
                    if (severity.equals("Minor")){
                        if(sever.equals("Minor") && (cellValToCompare.startsWith("PS") || cellValToCompare.startsWith("TS") || cellValToCompare.startsWith("Tech-Ops"))){

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
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableView.getItems().addAll(observableList);
                            caseCount++;
                            if (tableView.getItems().size() >= caseCount + 1) {
                                tableView.getItems().removeAll(observableList);
                            }
                        }
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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnRegCases.toFront();
                    lblStatus.setText("REGION VIEW - CASES BASED ON " + selectedRegion +  " REGION");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        }catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void accCaseUpdateTableView(String caseTableSelect, TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b, boolean bool) {
        int caseCount = 0;

        LocalDate dateToday = LocalDate.now();
        LocalDate caseUpdateDate = null;
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");


        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellValStat;
            HSSFCell cellVal1;
            HSSFCell cellVal2;

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
                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
            int accfiltnum = setAcc.size();

            for (int i = 0; i < accfiltnum; i++) {

                for (int k = 1; k < lastRow + 1; k++) {

                    cellVal = filtersheet.getRow(k).getCell(caseNextUpdateDateRef);
                    String cellValToCompare = cellVal.getStringCellValue();

                    cellValStat = filtersheet.getRow(k).getCell(caseStatRefCell);
                    String cellStat = cellValStat.getStringCellValue();

                    cellVal1 = filtersheet.getRow(k).getCell(caseAccountRef);
                    String acc = cellVal1.getStringCellValue();

                    ArrayList<String> array = new ArrayList<>();
                    ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                    if (!cellValToCompare.equals("NotSet")) {

                        caseUpdateDate = LocalDate.parse(cellValToCompare, formatter);
                    } else {
                        caseUpdateDate = null;
                    }

                    if (acc.equals(setAcc.get(i))) {

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays, age,
                                            caseUpdateDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays, age,
                                            caseUpdateDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays, age,
                                        caseUpdateDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            // Select and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }

                }
            });

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnAccounts.toFront();
                    lblStatus.setText("ACCOUNT(S) VIEW ( " + txAccounts.getText() +  ")");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void regCaseUpdateTableView(String caseTableSelect, TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b, boolean bool) {

        int caseCount = 0;

        LocalDate dateToday = LocalDate.now();
        LocalDate caseUpdateDate = null;
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");


        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellValStat;
            HSSFCell cellVal1;
            HSSFCell cellVal2;

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
                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            for (int k = 1; k < lastRow + 1; k++) {

                cellVal = filtersheet.getRow(k).getCell(caseNextUpdateDateRef);
                String cellValToCompare = cellVal.getStringCellValue();

                cellValStat = filtersheet.getRow(k).getCell(caseStatRefCell);
                String cellStat = cellValStat.getStringCellValue();

                cellVal1 = filtersheet.getRow(k).getCell(caseRegionRef);
                String region = cellVal1.getStringCellValue();

                ArrayList<String> array = new ArrayList<>();
                ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                if (!cellValToCompare.equals("NotSet")) {

                    caseUpdateDate = LocalDate.parse(cellValToCompare, formatter);
                } else {
                    caseUpdateDate = null;
                }

                if (region.equals(selectedRegion)) {

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
                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays, age,
                                        caseUpdateDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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

                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays, age,
                                        caseUpdateDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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

                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    caseUpdateDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableCases.getItems().addAll(observableList);
                            caseCount++;
                            if (tableCases.getItems().size() >= caseCount + 1) {
                                tableCases.getItems().removeAll(observableList);
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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            // Select and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }

                }
            });

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnRegCases.toFront();
                    lblStatus.setText("REGION VIEW - CASES BASED ON " + selectedRegion +  " REGION");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void accDueFilterView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, int ageDue, Boolean due) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
            int accfiltnum = setAcc.size();

            for (int i = 0; i <accfiltnum ; i++) {

                for (int k = 1; k < lastRow + 1; k++) {
                    cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                    String cellValToCompare = cellVal.getStringCellValue();
                    cellVal2 = filtersheet.getRow(k).getCell(caseAgeRefCell);
                    String caseAge = cellVal2.getStringCellValue();
                    cellVal3 = filtersheet.getRow(k).getCell(caseStatRefCell);
                    String caseStatus = cellVal3.getStringCellValue();
                    int ageCase = Integer.parseInt(caseAge);
                    cellVal4 = filtersheet.getRow(k).getCell(caseAccountRef);
                    String acc = cellVal4.getStringCellValue();
                    cellVal5 = filtersheet.getRow(k).getCell(rbbnDaysRefCell);
                    String ribdays = cellVal5.getStringCellValue();
                    int ribDays = Integer.parseInt(ribdays);

                    if (acc.equals(setAcc.get(i))) {

                        if (due) {

                            if ((cellValToCompare.equals(filter) && ribDays <= ageDue) && ((!caseStatus.equals("Develop Solution") &&
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
                                        array.get(3), array.get(4), array.get(5), ribDays, ageCase,
                                        localDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

                                tableCases.getItems().addAll(observableList);
                                caseCount++;
                                if (tableCases.getItems().size() >= caseCount + 1) {
                                    tableCases.getItems().removeAll(observableList);
                                }
                            }
                        } else {
                            if ((cellValToCompare.equals(filter) && ribDays > ageDue) && ((!caseStatus.equals("Develop Solution") &&
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
                                        array.get(3), array.get(4), array.get(5), ribDays, ageCase,
                                        localDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);


            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnAccounts.toFront();
                    lblStatus.setText("ACCOUNT(S) VIEW (" +txAccounts.getText() + ")");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void regDueFilterView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, int ageDue, Boolean due) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRegRefCell = i;
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
                cellVal4 = filtersheet.getRow(k).getCell(caseRegionRef);
                String region = cellVal4.getStringCellValue();
                cellVal5 = filtersheet.getRow(k).getCell(rbnDaysRegRefCell);
                String rbDays = cellVal5.getStringCellValue();
                int rbbnDays = Integer.parseInt(rbDays);

                if (region.equals(selectedRegion)) {

                    if (due) {

                        if ((cellValToCompare.equals(filter) && rbbnDays <= ageDue) && ((!caseStatus.equals("Develop Solution") &&
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
                                    array.get(3), array.get(4), array.get(5), rbbnDays, ageCase,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableCases.getItems().addAll(observableList);
                            caseCount++;
                            if (tableCases.getItems().size() >= caseCount + 1) {
                                tableCases.getItems().removeAll(observableList);
                            }
                        }
                    } else {
                        if ((cellValToCompare.equals(filter) && rbbnDays > ageDue) && ((!caseStatus.equals("Develop Solution") &&
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
                                    array.get(3), array.get(4), array.get(5), rbbnDays, ageCase,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableCases.getItems().addAll(observableList);
                            caseCount++;
                            if (tableCases.getItems().size() >= caseCount + 1) {
                                tableCases.getItems().removeAll(observableList);
                            }
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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);


            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnRegCases.toFront();
                    lblStatus.setText("REGION VIEW - CASES BASED ON " + selectedRegion +  " REGION");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void accInactiveTable(String columnSelect1, String filter1, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {
        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
            int accfiltnum = setAcc.size();

            for (int i = 0; i <accfiltnum ; i++) {

                for (int k = 1; k < lastRow + 1; k++) {
                    cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                    String cellValToCompare = cellVal.getStringCellValue();
                    cellVal2 = filtersheet.getRow(k).getCell(caseCellRef2);
                    String cellValToCompare2 = cellVal2.getStringCellValue();
                    cellVal3 = filtersheet.getRow(k).getCell(caseStatRefCell);
                    String caseStatus = cellVal3.getStringCellValue();
                    cellVal4 = filtersheet.getRow(k).getCell(caseAccountRef);
                    String acc = cellVal4.getStringCellValue();

                    if (acc.equals(setAcc.get(i))) {

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

                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableCases.getItems().addAll(observableList);
                            caseCount++;
                            if (tableCases.getItems().size() >= caseCount + 1) {
                                tableCases.getItems().removeAll(observableList);
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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnAccounts.toFront();
                    lblStatus.setText("ACCOUNT(S) VIEW (" + txAccounts.getText() +  ")");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void regInactiveTable(String columnSelect1, String filter1, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();
                cellVal2 = filtersheet.getRow(k).getCell(caseCellRef2);
                String cellValToCompare2 = cellVal2.getStringCellValue();
                cellVal3 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal3.getStringCellValue();
                cellVal4 = filtersheet.getRow(k).getCell(caseRegionRef);
                String region = cellVal4.getStringCellValue();

                if (region.equals(selectedRegion)) {

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
                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays, age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

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
                menu.getItems().add(openCaseDetails);
                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnRegCases.toFront();
                        lblStatus.setText("REGION VIEW - CASES BASED ON " + selectedRegion +  " REGION");
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });
            }
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void accEngineeringTableView(String columnSelect, String filter1, TableView<CaseTableView> tableCases, AnchorPane apnTableView){

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
            int accfiltnum = setAcc.size();

            for (int i = 0; i < accfiltnum; i++) {

                for (int k = 1; k < lastRow + 1; k++) {
                    cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                    String cellValToCompare = cellVal.getStringCellValue();

                    cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                    String caseStatus = cellVal2.getStringCellValue();

                    cellVal3 = filtersheet.getRow(k).getCell(caseAccountRef);
                    String acc = cellVal3.getStringCellValue();

                    if (acc.equals(setAcc.get(i))) {

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

                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableCases.getItems().addAll(observableList);
                            caseCount++;
                            if (tableCases.getItems().size() >= caseCount + 1) {
                                tableCases.getItems().removeAll(observableList);
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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnAccounts.toFront();
                    lblStatus.setText("ACCOUNT(S) VIEW (" + txAccounts.getText() +  ")");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void regEngineeringTableView(String columnSelect, String filter1, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();

                cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal2.getStringCellValue();

                cellVal3 = filtersheet.getRow(k).getCell(caseRegionRef);
                String region = cellVal3.getStringCellValue();

                if (region.equals(selectedRegion)) {

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

                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays, age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

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
                menu.getItems().add(openCaseDetails);
                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnRegCases.toFront();
                        lblStatus.setText("REGION VIEW - CASES BASED ON " + selectedRegion +  " REGION");
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });
            }
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void accTwoFilterTableView(String columnSelect1, String columnSelect2, String filter1, String filter2, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
            }

            ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
            int accfiltnum = setAcc.size();

            for (int i = 0; i <accfiltnum ; i++) {

                for (int k = 1; k < lastRow + 1; k++) {

                    cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                    String cellValToCompare = cellVal.getStringCellValue();
                    cellVal2 = filtersheet.getRow(k).getCell(caseCellRef2);
                    String cellValToCompare2 = cellVal2.getStringCellValue();

                    cellVal4 = filtersheet.getRow(k).getCell(caseStatRefCell);
                    String caseStatus = cellVal4.getStringCellValue();

                    cellVal3 = filtersheet.getRow(k).getCell(caseAccountRef);
                    String acc = cellVal3.getStringCellValue();

                    if (acc.equals(setAcc.get(i))) {

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

                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableCases.getItems().addAll(observableList);
                            caseCount++;
                            if (tableCases.getItems().size() >= caseCount + 1) {
                                tableCases.getItems().removeAll(observableList);
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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnAccounts.toFront();
                    lblStatus.setText("ACCOUNT(S) VIEW (" + txAccounts.getText() +  ")");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void regTwoFilterTableView(String columnSelect1, String columnSelect2, String filter1, String filter2, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }

            }
            for (int k = 1; k < lastRow + 1; k++) {

                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();
                cellVal2 = filtersheet.getRow(k).getCell(caseCellRef2);
                String cellValToCompare2 = cellVal2.getStringCellValue();

                cellVal4 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal4.getStringCellValue();

                cellVal3 = filtersheet.getRow(k).getCell(caseRegionRef);
                String region = cellVal3.getStringCellValue();

                if (region.equals(selectedRegion)) {

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

                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays, age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

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
                menu.getItems().add(openCaseDetails);
                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnRegCases.toFront();
                        lblStatus.setText("REGION VIEW - CASES BASED ON " + selectedRegion +  " REGION");
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });
            }
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void accWIPCaseTableView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal5;


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
                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }

            }

            ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
            int accfiltnum = setAcc.size();

            for (int i = 0; i <accfiltnum ; i++) {

                for (int k = 1; k < lastRow + 1; k++) {
                    cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                    String cellValToCompare = cellVal.getStringCellValue();

                    cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                    String caseStatus = cellVal2.getStringCellValue();

                    cellVal3 = filtersheet.getRow(k).getCell(caseAccountRef);
                    String acc = cellVal3.getStringCellValue();

                    if (acc.equals(setAcc.get(i))) {

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

                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableCases.getItems().addAll(observableList);
                            caseCount++;
                            if (tableCases.getItems().size() >= caseCount + 1) {
                                tableCases.getItems().removeAll(observableList);
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
            menu.getItems().add(openCaseDetails);

            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnAccounts.toFront();
                    lblStatus.setText("ACCOUNT(S) VIEW (" +txAccounts.getText() + ")");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }

    }

    private void regionWIPCaseTableView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal5;


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
                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }

            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();

                cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal2.getStringCellValue();

                cellVal3 = filtersheet.getRow(k).getCell(caseRegionRef);
                String region = cellVal3.getStringCellValue();

                if (region.equals(selectedRegion)) {

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

                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays, age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

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
                menu.getItems().add(openCaseDetails);

                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnRegCases.toFront();
                        lblStatus.setText("REGION VIEW - CASES BASED ON " + selectedRegion +  " REGION");
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });
            }
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void accOverviewWOHView(Boolean bool){

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal5;


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
                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }

            }

            ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
            int accfiltnum = setAcc.size();

            for (int j = 0; j < accfiltnum; j++) {

                for (int k = 1; k < lastRow + 1; k++) {
                    cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                    String cellValToCompare = cellVal.getStringCellValue();
                    cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                    String caseStatus = cellVal2.getStringCellValue();
                    cellVal3 = filtersheet.getRow(k).getCell(caseAccountRef);
                    String acc = cellVal3.getStringCellValue();

                    if (acc.equals(setAcc.get(j))) {

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

                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays, age,
                                        localDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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

                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays, age,
                                        localDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnAccounts.toFront();
                    lblStatus.setText("ACCOUNTS(S) VIEW - " + txAccounts.getText() + ")");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });

        }catch (Exception e){
            e.printStackTrace();
        }
    }

    private void regOverviewWOHView(Boolean bool) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();
                cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal2.getStringCellValue();
                cellVal3 = filtersheet.getRow(k).getCell(caseRegionRef);
                String region = cellVal3.getStringCellValue();

                if (region.equals(selectedRegion)) {

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

                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

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

                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

                            tableCases.getItems().addAll(observableList);
                            caseCount++;
                            if (tableCases.getItems().size() >= caseCount + 1) {
                                tableCases.getItems().removeAll(observableList);
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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            btnBack.setVisible(true);
            btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    apnRegCases.toFront();
                    lblStatus.setText("REGION VIEW - CASES BASED ON " + selectedRegion +  " REGION");
                    btnBack.setVisible(false);
                    btnToExcel.setVisible(false);
                    tableCases.getItems().clear();
                }
            });
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void regionOneFilterTableView(String columnSelect, String filter1, TableView tableCases, AnchorPane apnTableView, Boolean bool) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }

            }
            for (int k = 1; k < lastRow + 1; k++) {
                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();
                cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                String caseStatus = cellVal2.getStringCellValue();
                cellVal3 = filtersheet.getRow(k).getCell(caseRegionRef);
                String region = cellVal3.getStringCellValue();

                if (region.equals(selectedRegion)){

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

                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

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

                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                    localDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

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
                menu.getItems().add(openCaseDetails);
                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnRegCases.toFront();
                        lblStatus.setText("REGION VIEW - CASES BASED ON " + selectedRegion +  " REGION");
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });
            }
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void accOneFilterTableView(String columnSelect, String filter1, TableView tableCases, AnchorPane apnTableView, Boolean bool) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }

            }

            ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
            int accfiltnum = setAcc.size();

            for (int j = 0; j < accfiltnum; j++) {

                for (int k = 1; k < lastRow + 1; k++) {
                    cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                    String cellValToCompare = cellVal.getStringCellValue();
                    cellVal2 = filtersheet.getRow(k).getCell(caseStatRefCell);
                    String caseStatus = cellVal2.getStringCellValue();
                    cellVal3 = filtersheet.getRow(k).getCell(caseAccountRef);
                    String acc = cellVal3.getStringCellValue();


                    if (acc.equals(setAcc.get(j))) {
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
                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays, age,
                                        localDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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
                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays, age,
                                        localDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

                                tableCases.getItems().addAll(observableList);
                                caseCount++;
                                if (tableCases.getItems().size() >= caseCount + 1) {
                                    tableCases.getItems().removeAll(observableList);
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
                menu.getItems().add(openCaseDetails);
                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str=" + getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        } catch (Exception e) {
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
                        }
                    }
                });

                btnBack.setVisible(true);
                btnBack.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnAccounts.toFront();

                        ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
                        lblStatus.setText("ACCOUNT(S) VIEW (" +txAccounts.getText()+")" );
                        btnBack.setVisible(false);
                        btnToExcel.setVisible(false);
                        tableCases.getItems().clear();
                    }
                });

            }
        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    @FXML
    private void handleClicks(ActionEvent event) throws IOException, InvalidFormatException {

        if (event.getSource() == btnHome) {
            lblStatus.setText("GENERAL OVERVIEW");
            btnToExcel.setVisible(false);
            btnBack.setVisible(false);
            apnHome.toFront();
            apnHome.toFront();
            downData.parseData();
            downData.rectifyAccountNames();
            overviewPage();
        }

        if (event.getSource() == btnCases) {
            lblStatus.setText("MY CASES");
            btnToExcel.setVisible(false);
            btnBack.setVisible(false);
            myCasesPage();
            apnMyCases.toFront();
        }

        if (event.getSource() == btnProducts) {
            if (reg_product){
                lblStatus.setText("PRODUCT VIEW - REGION : " + regChoice.getValue());
            }else{
                lblStatus.setText("PRODUCT VIEW");
            }
            btnToExcel.setVisible(false);
            btnBack.setVisible(false);
            myProductsPage();
            apnProduct.toFront();
        }

        if (event.getSource() == btnProjects) {

            apnProjects.toFront();
            lblStatus.setText("PROJECTS VIEW");
            btnToExcel.setVisible(true);
            btnBack.setVisible(false);
            pnPrjNotes.setVisible(false);
            btnPrjNewNote.setVisible(false);
            btnPrjDelNote.setVisible(false);
            txtPrjNoteView.setVisible(false);
            downData.parseProjectData();
            projectsPage();
            initProjectTable();
            tableProjects.getItems().clear();
            tableProjects.setVisible(true);
            tableProjects.toFront();
            buildTableProjects(tableProjects, "All");

        }

        if (event.getSource() == btnCustomers) {
            lblStatus.setText("ACCOUNT(S) VIEW");
            apnAccounts.toFront();
            btnToExcel.setVisible(false);
            btnBack.setVisible(false);
            accountsPage();
        }

        if (event.getSource() == btnCustomerLoad) {
            if (!customerText.getText().isEmpty()) {
                customerViewPage();
            }
            tableCustomers.setVisible(false);
        }

        if (event.getSource() == btnMyNotes) {

            pnCaseDetailsNote.setVisible(false);
            caseNoteTable();

            if (caseNoteList.getItems().size()> 0) {
                lblStatus.setText("MY PERSONAL MEMO BOOK");
                apnNotes.toFront();
                btnToExcel.setVisible(false);
                btnBack.setVisible(false);
                txtShowCaseNotes.clear();
                caseNoteTable();
            }else {
                //alertNoNoteUser();
                //alertUser(strNoNote);
            }
        }

        if (event.getSource() == btnSettings) {
            lblStatus.setText("SETTINGS");
            btnToExcel.setVisible(false);
            btnBack.setVisible(false);
            apnSettings.toFront();
        }

        if (event.getSource() == btnProjection){

            apnProjection.toFront();
            projectionPage();
        }
        if (event.getSource() == btnSkillSet){
            rdEngMyTeam.setSelected(false);
            rdSkilMyTeam.setSelected(false);
            apnSkills.toFront();
            lblStatus.setText("SKILL SETS");
            skillEngSave();
            readAllUsers();
            readUsers();
            failSafeUsers();
        }

        if (event.getSource() == btnLogin) {
            btnToExcel.setVisible(false);
            btnBack.setVisible(false);
            //Connect to OKTA SSO
            connectOkta();
        }

        if (event.getSource() == btnLoadData) {
            //Download the related reports to work on them
            downData.downloadCSV();
        }

        if (event.getSource() == btnE1Cases) {
            if (e1Cases != 0) {
                lblStatus.setText("CRITICAL (OUTAGE) CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter1 = "Critical";
                initTableView(tableCases);
                oneFilterTableView(columnSelect, filter1, tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (e1Cases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (e2Cases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (outFollow == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (escCase == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (bcCases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (hotlist == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnWOH) {
            if (wohCases != 0) {
                lblStatus.setText("ACTIVE WORK ON HAND CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                overviewWOHView(true);
                tableCases.scrollToColumnIndex(0);
            }
            if (wohCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnInactive) {
            if (inactiveCases != 0) {
                lblStatus.setText("INACTIVE(PC/FA) CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                overviewWOHView(false);
                tableCases.scrollToColumnIndex(0);
            }
            if (inactiveCases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (BCwip == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (custActBC == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (custRpdBC == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (BCds == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (BCpc == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMJWIP) {
            if (MJwip != 0) {
                lblStatus.setText("MAJOR WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                lblStatus.setText("MAJOR CASE WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                overviewWIPCaseTableView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (MJwip == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNWIP) {
            if (MNwip != 0) {
                lblStatus.setText("MINOR WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                lblStatus.setText("MINOR CASE WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                overviewWIPCaseTableView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (MNwip == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (custActMJ == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (custRpdMJ == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (MJds == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (MJpc == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnBCDue) {
            if (bcDue != 0) {
                lblStatus.setText("BUSINESS CRITICAL DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 5, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (bcDue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnBCMissed) {
            if (misBCdue != 0) {
                lblStatus.setText("BUSINESS CRITICAL CASES MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 5, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (misBCdue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMJDue) {
            if (dueMJday != 0) {
                lblStatus.setText("MAJOR DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (dueMJday == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMJMissed) {
            if (misMJdue != 0) {
                lblStatus.setText("MAJOR CASES MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (misMJdue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNWac) {
            if (custActMN != 0) {
                lblStatus.setText("MINOR CASES PENDING INFORMATION FROM CUSTOMER");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Minor";
                String filter2 = "Customer action";
                initTableView(tableCases);
                twoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (custActMN == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNupdated) {
            if (custRpdMN != 0) {
                lblStatus.setText("MINOR CASES PENDING OWNER ACTION");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String columnSelect2 = "Currently Responsible";
                String filter1 = "Minor";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterTableView(columnSelect1, columnSelect2, filter1, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (custRpdMN == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNEngineering) {
            if (MNds != 0) {
                lblStatus.setText("MINOR CASES WITH DESIGN");
                tableCases.getItems().clear();
                tableCases.getItems().clear();
                String columSelect = "Severity";
                String filter1 = "Minor";
                initTableView(tableCases);
                overviewEngineeringTableView(columSelect, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (MNds == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNINACT) {
            if (MNpc != 0) {
                lblStatus.setText("MINOR CASES PENDING CLOSURE");
                tableCases.getItems().clear();
                String columnSelect1 = "Severity";
                String filter1 = "Minor";
                initTableView(tableCases);
                overViewInactiveTable(columnSelect1, filter1, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (MNpc == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNDue) {
            if (dueMNday != 0) {
                lblStatus.setText("MINOR DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (dueMJday == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNMissed) {
            if (misMNdue != 0) {
                lblStatus.setText("MINOR CASES MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                overviewDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (misMNdue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnTSQueue) {
            if (queueTS != 0) {
                lblStatus.setText("CASES IN RTS QUEUE");
                tableCases.getItems().clear();
                String columnselect = "Case Owner";
                String filter = "TS";
                initTableView(tableCases);
                overviewQueueView(columnselect, filter, tableCases, apnTableView, "RTS QUEUE");
                tableCases.scrollToColumnIndex(0);
            }
            if (queueTS == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnPSQueue) {
            if (queuePS != 0) {
                lblStatus.setText("CASES IN GPS QUEUE");
                tableCases.getItems().clear();
                String e2TableSelect = "Case Owner";
                String e2TableSelect2 = "PS";
                initTableView(tableCases);
                overviewQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView, "GPS QUEUE");
                tableCases.scrollToColumnIndex(0);
            }
            if (queuePS == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnBCQueue) {
            if (bcQueue != 0) {
                lblStatus.setText("BC CASES IN QUEUE");
                tableCases.getItems().clear();
                String e2TableSelect = "Severity";
                String e2TableSelect2 = "Business Critical";
                initTableView(tableCases);
                overviewSevQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (bcQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMJQueue) {
            if (mjQueue != 0) {
                lblStatus.setText("MAJOR CASES IN QUEUE");
                tableCases.getItems().clear();
                String e2TableSelect = "Severity";
                String e2TableSelect2 = "Major";
                initTableView(tableCases);
                overviewSevQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (mjQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNQueue) {
            if (mnQueue != 0) {
                lblStatus.setText("MINOR CASES IN QUEUE");
                tableCases.getItems().clear();
                String e2TableSelect = "Severity";
                String e2TableSelect2 = "Minor";
                initTableView(tableCases);
                overviewSevQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (mnQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnUpdateToday) {
            if (updateToday != 0) {
                lblStatus.setText("NEXT CASE UPDATE TODAY LIST");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                caseUpdateTableView(columnSelect, tableCases, apnTableView, true, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (updateToday == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnUpdateMissed) {
            if (updateMissed != 0) {
                lblStatus.setText("NEXT CASE UPDATE MISSED LIST");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                caseUpdateTableView(columnSelect, tableCases, apnTableView, false, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (updateMissed == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnUpdateNull) {
            if (updateNull != 0) {
                lblStatus.setText("NEXT CASE UPDATE FIELD NOT SET CASE LIST");
                tableCases.getItems().clear();
                String columnSelect = "Next Case Update";
                initTableView(tableCases);
                caseUpdateTableView(columnSelect, tableCases, apnTableView, false, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (updateNull == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myE1Case == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myE2Cases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myOutFollow == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myEscCases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myBCCases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myHotList == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyWOH) {
            if (myWOHCases != 0) {
                lblStatus.setText("MY WORK ON HAND CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                myWOHTableView(tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (myWOHCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyInactive) {
            if (myInactiveCases != 0) {
                lblStatus.setText("MY INACTIVE (PENDING CLOSURE) CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                myWOHTableView(tableCases, apnTableView, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (myInactiveCases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myBCWIP == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myBCWac == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myBCupdated == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myBCDSCases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myBCInactiveCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyBCQueue) {
            if (myBCInQueue != 0) {
                lblStatus.setText("BC CASES IN MY QUEUE(S)");
                tableCases.getItems().clear();
                String e2TableSelect = "Severity";
                String e2TableSelect2 = "Business Critical";
                initTableView(tableCases);
                mySevQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (myBCInQueue == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myMJWIP == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnMyMNWIP) {

            if (myMNWIP != 0) {

                lblStatus.setText("MY MINOR WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                String columFilter = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                overviewMyWIPCaseTableView(columFilter, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (myMNWIP == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myMJWAC == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMNWac) {

            if (myMNWAC != 0) {
                lblStatus.setText("MY MINOR CASES PENDING INFORMATION FROM CUSTOMER");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Minor";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer action";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (myMNWAC == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMJupdated) {

            if (myMJUpdated != 0) {
                lblStatus.setText("MY MAJOR CASES CUSTOMER PROVIDED UPDATE");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Major";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (myMJUpdated == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMNupdated) {

            if (myMNUpdated != 0) {
                lblStatus.setText("MY MINOR CASES CUSTOMER PROVIDED UPDATE");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Minor";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (myMNUpdated == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }

            if (myMJDSCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMNEngineering) {

            if (myMNEng != 0) {

                lblStatus.setText("MY MINOR CASES WITH DESIGN");
                tableCases.getItems().clear();
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Minor";
                String columSelect2 = "Status";
                String filter2 = "Develop Solution";
                initTableView(tableCases);
                twoFilterMyTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }

            if (myMNEng == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (myMJInactiveCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMNINACT) {

            if (myMNINAct != 0) {

                lblStatus.setText("MY MINOR INACTIVE (PENDING CLOSURE) CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                inactiveCasesMyTableView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (myMNINAct == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMJQueue) {
            if (myMJInQueue != 0) {
                lblStatus.setText("MAJOR CASES IN MY QUEUE(S)");
                tableCases.getItems().clear();
                String e2TableSelect = "Severity";
                String e2TableSelect2 = "Major";
                initTableView(tableCases);
                mySevQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (myMJInQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMNQueue) {
            if (myMNInQueue != 0) {
                lblStatus.setText("MINOR CASES IN MY QUEUE(S)");
                tableCases.getItems().clear();
                String e2TableSelect = "Severity";
                String e2TableSelect2 = "Minor";
                initTableView(tableCases);
                mySevQueueView(e2TableSelect, e2TableSelect2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (myMNInQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyBCDue) {

            if (myBCDueCases != 0) {
                lblStatus.setText("MY BUSINESS CRITICAL DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                myDueFilterView(columnSelect, filter, tableCases, apnTableView, 5, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (myBCDueCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyBCMissed) {

            if (myBCMissedCases != 0) {
                lblStatus.setText("MY BUSINESS CRITICAL CASED MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                myDueFilterView(columnSelect, filter, tableCases, apnTableView, 5, false);
                tableCases.scrollToColumnIndex(0);
            }

            if (myBCMissedCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMJDue) {

            if (myMJDueCases != 0) {
                lblStatus.setText("MY MAJOR DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                myDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, true);
                tableCases.scrollToColumnIndex(0);
            }

            if (myMJDueCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMNDue) {

            if (myMNDue != 0) {
                lblStatus.setText("MY MINOR DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                myDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, true);
                tableCases.scrollToColumnIndex(0);
            }

            if (myMNDue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMJMissed) {

            if (myMJMissedCases != 0) {
                lblStatus.setText("MY MAJOR CASED MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                myDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, false);
                tableCases.scrollToColumnIndex(0);
            }

            if (myMJMissedCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyMNMissed) {

            if (myMNMissed != 0) {
                lblStatus.setText("MY MINOR CASED MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                myDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, false);
                tableCases.scrollToColumnIndex(0);
            }

            if (myMNMissed == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyQueue) {

            if (myQueuedCases != 0) {
                lblStatus.setText("CASES IN MY QUEUE(S)");
                tableCases.getItems().clear();
                String columnSelect = "Case Owner";
                initTableView(tableCases);
                createMyQueueCaseView(columnSelect, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (myQueuedCases == 0) {
                alertUser(strAlert);
            }
        }
/*        if (event.getSource() == btnMyCoOwnQueue) {

            if (myCoOwnerQueueCases != 0) {
                lblStatus.setText("CASES IN MY CO-OWNER QUEUE(S)");
                tableCases.getItems().clear();
                initTableView(tableCases);
                createMyCoOwnerQueueCaseView(false);
                tableCases.scrollToColumnIndex(0);
            }
            if (myCoOwnerQueueCases == 0) {
                alertUser(strAlert);
            }
        }*/

/*        if (event.getSource() == btnMyCoQueueAssigned){
            if (myCoOwnerQueueCasesAssigned != 0) {
                lblStatus.setText("CASES ASSIGNED FROM MY CO-OWNER QUEUE(S)");
                tableCases.getItems().clear();
                initTableView(tableCases);
                createMyCoOwnerQueueCaseView(true);
                tableCases.scrollToColumnIndex(0);
            }
            if (myCoOwnerQueueCasesAssigned == 0){
                alertUser(strAlert);
            }
        }*/

        if (event.getSource() == btnMyUpdateToday) {

            if (myUpdateToday != 0) {

                lblStatus.setText("NEXT CASE UPDATE TODAY LIST");
                tableCases.getItems().clear();
                String caseTableSelect = "Next Case Update";
                initTableView(tableCases);
                mycaseUpdateTableView(caseTableSelect, tableCases, apnTableView, true, true);
                tableCases.scrollToColumnIndex(0);
            }

            if (myUpdateToday == 0) {
                alertUser(strAlert);
            }

        }

        if (event.getSource() == btnMyUpdateMissed) {

            if (myUpdateMissed != 0) {

                lblStatus.setText("NEXT CASE UPDATE MISSED LIST");
                tableCases.getItems().clear();
                String caseTableSelect = "Next Case Update";
                initTableView(tableCases);
                mycaseUpdateTableView(caseTableSelect, tableCases, apnTableView, false, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (myUpdateMissed == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMyUpdateNull) {
            if (myUpdateNull != 0) {

                lblStatus.setText("NEXT CASE UPDATE FIELD NOT SET CASE LIST");
                tableCases.getItems().clear();
                String caseTableSelect = "Next Case Update";
                initTableView(tableCases);
                mycaseUpdateTableView(caseTableSelect, tableCases, apnTableView, false, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (myUpdateNull == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }

            if (prodE1Case == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodE2Cases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodOutFollow == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodEscCases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodBCCases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodHotList == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnWOHProd) {
            if (prodWOHCases != 0) {
                lblStatus.setText("PRODUCT VIEW - ACTIVE WORK ON HAND");
                tableCases.getItems().clear();
                initTableView(tableCases);
                prodWOHTable(tableCases, apnTableView, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodWOHCases == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnInactiveProd) {
            if (prodInactiveCases != 0) {
                lblStatus.setText("PRODUCT VIEW - INACTIVE WORK ON HAND");
                tableCases.getItems().clear();
                initTableView(tableCases);
                prodWOHTable(tableCases, apnTableView, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodInactiveCases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodBCWIP == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodBCWac == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodBCupdated == 0) {
                alertUser(strAlert);
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
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodBCDSCases == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodBCInactiveCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMJWIPProd) {
            if (prodMJWIP != 0) {
                lblStatus.setText("PRODUCT VIEW - MAJOR WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                String columnSelect = "Severity";
                String filter = "Major";
                productWIPCaseView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMJWIP == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNWIPProd) {
            if (prodMNWIP != 0) {
                lblStatus.setText("PRODUCT VIEW - MINOR WORK IN PROGRESS CASES");
                tableCases.getItems().clear();
                initTableView(tableCases);
                String columnSelect = "Severity";
                String filter = "Minor";
                productWIPCaseView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMNWIP == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMJWacProd) {
            if (prodMJWAC != 0) {
                lblStatus.setText("PRODUCT VIEW - MAJOR CASES PENDING CUSTOMER ACTION");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Major";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer action";
                initTableView(tableCases);
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMJWAC == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNWacProd) {
            if (prodMNWAC != 0) {
                lblStatus.setText("PRODUCT VIEW - MINOR CASES PENDING CUSTOMER ACTION");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Minor";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer action";
                initTableView(tableCases);
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMNWAC == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMJupdatedProd) {
            if (prodMJUpdated != 0) {
                lblStatus.setText("PRODUCT VIEW - MAJOR CASES CUSTOMER PROVIDED UPDATE");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Major";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMJUpdated == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnMNupdatedProd) {
            if (prodMNUpdated != 0) {
                lblStatus.setText("PRODUCT VIEW - MINOR CASES CUSTOMER PROVIDED UPDATE");
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Minor";
                String columSelect2 = "Currently Responsible";
                String filter2 = "Customer updated";
                initTableView(tableCases);
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMNUpdated == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMJDSCases == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnMNEngineeringProd) {
            if (prodMNEng != 0) {
                lblStatus.setText("PRODUCT VIEW - MINOR CASES WITH DESIGN");
                tableCases.getItems().clear();
                tableCases.getItems().clear();
                String columSelect1 = "Severity";
                String filter1 = "Minor";
                String columSelect2 = "Status";
                String filter2 = "Develop Solution";
                initTableView(tableCases);
                twoFilterProductTableView(columSelect1, filter1, columSelect2, filter2, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMNEng == 0) {
                alertUser(strAlert);
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
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMJInactiveCases == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnMNINACTProd) {
            if (prodMNInact != 0) {

                lblStatus.setText("PRODUCT VIEW - MINOR INACTIVE (PC & FA) CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                inactiveCasesProductTableView(columnSelect, filter, tableCases, apnTableView);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMNInact == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnBCDueProd) {

            if (prodBCDueCases != 0) {
                lblStatus.setText("PRODUCT VIEW - BUSINESS CRITICAL DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                productDueFilterView(columnSelect, filter, tableCases, apnTableView, 5, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodBCDueCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnBCMissedProd) {

            if (prodBCMissedCases != 0) {
                lblStatus.setText("PRODUCT VIEW - BUSINESS CRITICAL CASES MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCases);
                productDueFilterView(columnSelect, filter, tableCases, apnTableView, 5, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodBCMissedCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMJDueProd) {

            if (prodMJDueCases != 0) {
                lblStatus.setText("PRODUCT VIEW - MAJOR DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                productDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMJDueCases == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnMNDueProd) {

            if (prodMNDue != 0) {
                lblStatus.setText("PRODUCT VIEW - MINOR DUE CASES");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                productDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, true);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMNDue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMJMissedProd) {

            if (prodMJMissedCases != 0) {
                lblStatus.setText("PRODUCT VIEW - MAJOR CASED MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Major";
                initTableView(tableCases);
                productDueFilterView(columnSelect, filter, tableCases, apnTableView, 15, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMJMissedCases == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNMissedProd) {

            if (prodMNMissed != 0) {
                lblStatus.setText("PRODUCT VIEW - MINOR CASED MISSED DUE");
                tableCases.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Minor";
                initTableView(tableCases);
                productDueFilterView(columnSelect, filter, tableCases, apnTableView, 30, false);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMNMissed == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnTSQueueProd) {
            if (prodQueueTS != 0) {
                lblStatus.setText("PRODUCT VIEW - CASES IN RTS QUEUE");
                tableCases.getItems().clear();
                initTableView(tableCases);
                String sev = "All";
                productViewCasesQueued(tableCases, apnTableView, false, sev);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodQueueTS == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnBCQueueProd) {
            if (prodBCQueue != 0) {
                lblStatus.setText("PRODUCT VIEW - BC CASES IN QUEUE(S)");
                tableCases.getItems().clear();
                initTableView(tableCases);
                String sev = "BC";
                productViewCasesQueued(tableCases, apnTableView, false, sev);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodBCQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMJQueueProd) {
            if (prodMJQueue != 0) {
                lblStatus.setText("PRODUCT VIEW - MAJOR CASES IN QUEUE(S)");
                tableCases.getItems().clear();
                initTableView(tableCases);
                String sev = "Major";
                productViewCasesQueued(tableCases, apnTableView, false, sev);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMJQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnMNQueueProd) {
            if (prodMNQueue != 0) {
                lblStatus.setText("PRODUCT VIEW - MINOR CASES IN QUEUE(S)");
                tableCases.getItems().clear();
                initTableView(tableCases);
                String sev = "Minor";
                productViewCasesQueued(tableCases, apnTableView, false, sev);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodMNQueue == 0) {
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnPSQueueProd) {
            if (prodQueuePS != 0) {
                lblStatus.setText("PRODUCT VIEW - CASES IN GPS QUEUE");
                tableCases.getItems().clear();
                initTableView(tableCases);
                String sev = "All";
                productViewCasesQueued(tableCases, apnTableView, true, sev);
                tableCases.scrollToColumnIndex(0);
            }
            if (prodQueuePS == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnCustomerCritical) {
            if (customerE1 != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Critical";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, true);
                tableCustomers.scrollToColumnIndex(0);
            }
            if (customerE1 == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnCustomerE2) {
            if (customerE2 != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Severity";
                String filter = "E2";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, true);
                tableCustomers.scrollToColumnIndex(0);
            }
            if (customerE2 == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnCustomerOutFollow) {
            if (customerOutFol != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Outage Follow-Up";
                String filter = "1";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, true);
                tableCustomers.scrollToColumnIndex(0);
            }
            if (customerOutFol == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnCustomerEscalated) {
            if (customerEsc != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Escalated By";
                String filter = "NotSet";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, false);
                tableCustomers.scrollToColumnIndex(0);
            }
            if (customerEsc == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnCustomerHotIssues) {
            if (customerHot != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Support Hotlist Level";
                String filter = "NotSet";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, false);
                tableCustomers.scrollToColumnIndex(0);
            }
            if (customerHot == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnCustomerBC) {
            if (customerBC != 0) {
                tableCustomers.getItems().clear();
                String columnSelect = "Severity";
                String filter = "Business Critical";
                initTableView(tableCustomers);
                customerTable(columnSelect, filter, tableCustomers, true);
                tableCustomers.scrollToColumnIndex(0);
            }
            if (customerBC == 0) {
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnCustomerActWOH) {
            if (customerWoh != 0) {
                tableCustomers.getItems().clear();
                initTableView(tableCustomers);
                customerWOHTable(tableCustomers, true);
                tableCustomers.scrollToColumnIndex(0);
            }
            if (customerWoh == 0) {
                alertUser(strAlert);
                tableCustomers.setVisible(false);
            }
        }
        if (event.getSource() == btnMyRegion){
            apnRegCases.toFront();
            selectedRegion = regChoice.getSelectionModel().getSelectedItem();
            lblStatus.setText("REGION VIEW - CASES BASED ON " + selectedRegion +  " REGION");
            btnToExcel.setVisible(false);
            btnBack.setVisible(false);
            regionCases();
        }
        if (event.getSource() == checkRegProduct){
            if (checkRegProduct.isSelected()){
                reg_product = true;
            }else{
                reg_product = false;
            }
        }
    }

    private void caseNoteTable(){

        caseNoteList.getItems().clear();
        ObservableList<String> Notes = FXCollections.observableArrayList();

        ArrayList<String> details = new ArrayList<String>();
        File repo = new File(noteFolder);

        if (!repo.exists()){
            //alertNoNoteUser();
            String strNoNote = "THERE IS NO PERSONAL NOTE..." + "\n" + "\n" + "PLEASE CREATE PERSONAL NOTE FIRST!";
            alertUser(strNoNote);


        }else {

            File[] fileList = repo.listFiles();

            for (int i = 0; i < fileList.length; i++) {
                if (!fileList[i].getName().equals("Project")){
                    Notes.addAll(fileList[i].getName());
                }
            }

            caseNoteList.getItems().addAll(Notes);
            caseNoteList.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
            caseNoteList.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    details.clear();
                    txtShowCaseNotes.clear();

                    try {

                        if (caseNoteList.getItems().size() > 0) {

                            String selectedCase = caseNoteList.getSelectionModel().getSelectedItem().toString();
                            File caseDetails = new File(detailsFolder + "\\" + selectedCase);

                            ClipboardContent content = new ClipboardContent();
                            content.putString(selectedCase);
                            Clipboard.getSystemClipboard().setContent(content);

                            if (caseDetails.isFile()) {
                                Scanner s = null;
                                try {
                                    s = new Scanner(caseDetails);
                                } catch (Exception e) {
                                    logger.log(Level.WARNING, "Case Details - No File Found...: ", e);
                                }
                                while (s.hasNextLine()) {
                                    details.add(s.nextLine());
                                }
                                s.close();
                                txtCaseNumnberNote.setText(details.get(0));
                                txtCaseSeverityNote.setText(details.get(1));
                                txtCaseStatusNote.setText(details.get(2));
                                txtCaseOwnerNote.setText(details.get(3));
                                txtCaseCoOwnerNote.setText(details.get(4));
                                //txtCaseCoQueueNote.setText(details.get(5));
                                txtCaseAgeNote.setText(details.get(7));
                                txtCaseTypeNote.setText(details.get(12));
                                txtCaseProductNote.setText(details.get(13));
                                txtCaseSubjectNote.setText(details.get(14));
                                txtCaseAccountNote.setText(details.get(15));
                                txtCaseRegionNote.setText(details.get(16));

                                if (!details.get(9).equals("NotSet")) {
                                    checkBoxEscalatedNote.setSelected(true);
                                }
                                else{
                                    checkBoxEscalatedNote.setSelected(false);
                                }
                                if (!details.get(10).equals("NotSet")){
                                    checkBoxHotIssueNote.setSelected(true);
                                }else{
                                    checkBoxHotIssueNote.setSelected(false);
                                }
                            }

                            pnCaseDetailsNote.setVisible(true);
                            btnViewNote.setVisible(true);
                            btnViewComment.setVisible(true);
                            btnAddNewNote.setVisible(true);
                            btnDelNote.setVisible(true);

                            txtShowCaseNotes.clear();

                            File caseNote = new File(noteFolder + "\\" + selectedCase);

                            if (caseNote.isFile()) {
                                Scanner s = null;
                                try {
                                    s = new Scanner(caseNote);
                                } catch (Exception e) {
                                    logger.log(Level.WARNING, "Personal Memo - No File Found...: ", e);
                                }
                                while (s.hasNextLine()) {
                                    txtShowCaseNotes.appendText(s.nextLine() + "\n");
                                }
                                s.close();
                            }
                            spnNote.setVisible(true);

                            btnViewNote.setOnMouseClicked(new EventHandler<MouseEvent>() {
                                @Override
                                public void handle(MouseEvent event) {

                                    txtShowCaseNotes.clear();

                                    File caseNote = new File(noteFolder+ "\\" + selectedCase);

                                    if (caseNote.isFile()) {
                                        Scanner s = null;
                                        try {
                                            s = new Scanner(caseNote);
                                        } catch (Exception e) {
                                            logger.log(Level.WARNING, "Notes - No Case Selection File Found...: ", e);
                                        }
                                        while (s.hasNextLine()) {
                                            txtShowCaseNotes.appendText(s.nextLine() + "\n");
                                        }
                                        s.close();
                                    }
                                    spnNote.setVisible(true);

                                }
                            });

                            btnViewComment.setOnMouseClicked(new EventHandler<MouseEvent>() {
                                @Override
                                public void handle(MouseEvent event) {
                                    viewComments(details.get(0));
                                }
                            });

                            btnAddNewNote.setOnMouseClicked(new EventHandler<MouseEvent>() {
                                @Override
                                public void handle(MouseEvent event) {

                                    ClipboardContent content = new ClipboardContent();
                                    content.putString(selectedCase);
                                    Clipboard.getSystemClipboard().setContent(content);
                                    txtShowCaseNotes.clear();
                                    spnNote.setVisible(false);
                                    newCaseNote();
                                }
                            });

                            btnDelNote.setOnMouseClicked(new EventHandler<MouseEvent>() {
                                @Override
                                public void handle(MouseEvent event) {
                                    File caseNoteFile = new File(noteFolder + "\\" + selectedCase);
                                    File caseDetail = new File(detailsFolder + "\\" + selectedCase);

                                    caseNoteFile.delete();
                                    caseDetail.delete();
                                    pnCaseDetailsNote.setVisible(false);
                                    caseNoteList.getItems().remove(selectedCase);
                                    if (caseNoteList.getItems().size() == 0) {
                                        apnMyCases.toFront();
                                        lblStatus.setText("MY CASES");

                                    }
                                    txtShowCaseNotes.clear();
                                }
                            });
                        }

                    } catch (Exception e) {
                        logger.log(Level.WARNING, "No Personal Memo Found...: ", e);
                    }
                }
            });
        }
    }

    private void newCaseNote(){

        Parent root;
        try {
            root = FXMLLoader.load(getClass().getClassLoader().getResource("home/CaseNote.fxml"));
            Stage stage = new Stage();
            stage.setTitle("ADD PERSONAL CASE NOTE");
            stage.getIcons().add(new Image("home/image/rbbicon.png"));
            stage.setScene(new Scene(root, 650, 400));
            stage.show();
            stage.setMinWidth(650);
            stage.setMinHeight(820);
            stage.setMaxWidth(650);
            stage.setMaxHeight(820);

            //saveCaseDetails();

        }
        catch (IOException e) {
            logger.log(Level.WARNING, "Not Able To Open New Memeo Page...: ", e);
        }
    }

    private void projectionPage(){

        WebEngine project = projectWeb.getEngine();

        btnForOverAll.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                try {
                    project.load(String.valueOf(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Forecast\\ALL_Case_Arrival_Forecast_stats_arima.html").toURI().toURL()));
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        });

        btnForIMS.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                try {
                    project.load(String.valueOf(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Forecast\\IMS_Case_Arrival_Forecast_stats_arima.html").toURI().toURL()));
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        });

        btnForMM.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                try {
                    project.load(String.valueOf(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Forecast\\MM_Case_Arrival_Forecast_stats_arima.html").toURI().toURL()));
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        });

        /*btnForecastRun.setVisible(false);
        apnProjection.toFront();


        forecastAll.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                forecastProductSelect.clear();
                btnForecastRun.setVisible(true);
                apnForecastSel.setVisible(false);
                btnForecastRun.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        WebEngine project = projectWeb.getEngine();

                        try {
                            project.load(String.valueOf(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Forecast\\Forecast_in_Future_Overall.html").toURI().toURL()));
                        }catch (Exception e){
                            e.printStackTrace();
                        }
                    }
                });
            }
        });

        forecastProductSelect.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                if (forecastProductSelect.equals("")){
                    btnForecastRun.setVisible(false);
                }

                forecastAll.setSelected(false);
                apnForecastSel.setVisible(true);

                forecastSelect.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        apnForecastSel.setVisible(true);
                        lstForecast.getItems().clear();
                        ObservableList<String> availProd = FXCollections.observableArrayList();

                        try {
                            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_user_prod.xls")));
                            HSSFSheet sheet = workbook.getSheetAt(0);
                            HSSFCell cellVal;

                            int lastRow = sheet.getLastRowNum();
                            int cellnum = sheet.getRow(0).getLastCellNum();

                            for (int i = 0; i < cellnum; i++) {
                                String filterColName = sheet.getRow(0).getCell(i).toString();

                                if (filterColName.equals("Support Product")) {
                                    caseProductRef = i;
                                }
                            }

                            ArrayList<String> prodArray = new ArrayList<>();

                            for (int i = 1; i < lastRow; i++) {

                                cellVal = sheet.getRow(i).getCell(caseProductRef);
                                String productName = "";

                                if (cellVal != null) {
                                    productName = cellVal.getStringCellValue();
                                }
                                prodArray.add(productName);
                            }

                            prodArray = (ArrayList) prodArray.stream().distinct().collect(Collectors.toList());
                            Collections.sort(prodArray);

                            int size = prodArray.size();

                            for (int i = 0; i < size; i++) {
                                availProd.addAll(prodArray.get(i));
                            }
                            lstForecast.getItems().addAll(availProd);

                            FilteredList<String> filteredProduct = new FilteredList((ObservableList) availProd, p -> true);

                            forecastSelect.textProperty().addListener((observable, oldValue, newValue) -> {
                                filteredProduct.setPredicate(string -> {

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

                            lstForecast.setItems(filteredProduct);

                            lstForecast.getSelectionModel().selectedItemProperty().addListener((obs, newVal, oldVal) -> {

                                lstForecast.setOnMouseClicked(new EventHandler<MouseEvent>() {
                                    @Override
                                    public void handle(MouseEvent event) {

                                        if (event.getClickCount() > 1) {
                                            try {

                                                if (lstForecast.getSelectionModel().getSelectedItem() != null) {
                                                    String selectedProduct = lstForecast.getSelectionModel().getSelectedItem();
                                                    //filteredAccounts.add(selectedAcc.getAccountName());
                                                    forecastProductSelect.setText(selectedProduct);
                                                    btnForecastRun.setVisible(true);
                                                    apnForecastSel.setVisible(false);
                                                }
                                            } catch (Exception e) {
                                                logger.log(Level.WARNING, "Unable To Add Product to Selected By Click...", e);
                                            }
                                        }
                                    }
                                });
                            });

                        }catch (Exception e){
                            e.printStackTrace();
                        }
                    }
                });
            }
        });*/
    }

    private void projectionPage2(){

        String user = "santera";
        String password = "santera1";
        String host = "47.168.122.68";
        int port = 22;
        String command1="ls -ltr";

        try{

            java.util.Properties config = new java.util.Properties();

            JSch jsch = new JSch();
            Session session = jsch.getSession(user, host, port);
            session.setPassword(password);
            session.setConfig("PreferredAuthentications",
                    "publickey,keyboard-interactive,password");
            session.setConfig("StrictHostKeyChecking", "no");
            System.out.println("Establishing Connection...");
            session.connect();
            System.out.println("Connection established.");

            Channel channel=session.openChannel("exec");
            ((ChannelExec)channel).setCommand(command1);
            channel.setInputStream(null);
            ((ChannelExec)channel).setErrStream(System.err);

            channel.connect();

            /*InputStream in=channel.getInputStream();
            channel.connect();
            byte[] tmp=new byte[1024];
            while(true){
                while(in.available()>0){
                    int i=in.read(tmp, 0, 1024);
                    if(i<0)break;
                    System.out.print(new String(tmp, 0, i));
                }
                if(channel.isClosed()){
                    System.out.println("exit-status: "+channel.getExitStatus());
                    break;
                }
                try{Thread.sleep(1000);}catch(Exception ee){}
            }*/

            Session session2 = jsch.getSession(user, host);
            session2.setPassword(password);
            session2.setConfig("PreferredAuthentications",
                    "publickey,keyboard-interactive,password");
            session2.setConfig("StrictHostKeyChecking", "no");

            System.out.println("Establishing Connection...");
            session2.connect();
            System.out.println("Connection established.");

            ChannelSftp sftpChannel = (ChannelSftp) session.openChannel("sftp");
            sftpChannel.connect();

            File projectionDir = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Projection");

            if (!projectionDir.exists()) {

                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Projection").mkdir();
            }

            sftpChannel.get("/space/Santera/msg", System.getProperty("user.home") + "/Documents/CMT/Projection/msg");

            channel.disconnect();
            session.disconnect();
            sftpChannel.disconnect();
            session2.disconnect();
            System.out.println("DONE");


        }catch (Exception e){
            e.printStackTrace();
        }

    }

    private void skillEngSave(){

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
        Collections.sort(settingsUsers);

        try {

            FileWriter writer = new FileWriter(new File(skillSetFolder + "\\users.txt"));
            int size = settingsUsers.size();
            for (int i = 0; i < size; i++) {
                String str = settingsUsers.get(i);
                writer.write(str);
                if (i < size - 1)
                    writer.write("\n");
            }

            writer.close();

        } catch (Exception e) {
            logger.log(Level.WARNING, "Skill Set - Unable to Save users...: ", e);;
        }
    }

    private void saveCaseDetails(CaseTableView caseview) {

        /*TablePosition tablePosition = (TablePosition) tableCases.getSelectionModel().getSelectedCells().get(0);
        int row = tablePosition.getRow();
        CaseTableView caseview = (CaseTableView) tableCases.getItems().get(row);
        TableColumn tableColumn = tablePosition.getTableColumn();*/

        selectedCase = new ArrayList<>();
        selectedCase.add(caseview.getCaseNumber());
        selectedCase.add(caseview.getCaseSeverity());
        selectedCase.add(caseview.getCaseStatus());
        selectedCase.add(caseview.getCaseOwner());
        selectedCase.add(caseview.getCaseCoOwner());
        selectedCase.add(caseview.getCaseCoOwnerQueue());
        selectedCase.add(caseview.getCaseResponsible().toString());
        selectedCase.add(caseview.getCaseAge().toString());
        if(caseview.getNextCaseUpdate() != null){
            selectedCase.add(caseview.getNextCaseUpdate().toString());
        }else{
            selectedCase.add("NotSet");
        }
        selectedCase.add(caseview.getCaseEscalatedBy());
        selectedCase.add(caseview.getCaseHotList());
        selectedCase.add(caseview.getCaseOutFollow());
        selectedCase.add(caseview.getCaseSupportType());
        selectedCase.add(caseview.getCaseProduct());
        selectedCase.add(caseview.getCaseSubject());
        selectedCase.add(caseview.getCaseAccount());
        selectedCase.add(caseview.getCaseRegion());
        selectedCase.add(caseview.getCaseSecurity());

        int selectedsize= selectedCase.size();

        try {

            File caseDetailsFile = new File(detailsFolder + "\\" + caseview.getCaseNumber());

            File caseSelFile = new File(detailsFolder + "\\" + caseview.getCaseNumber());

            FileWriter writer = new FileWriter(caseSelFile);

            for (int i = 0; i < selectedsize; i++) {

                writer.write(selectedCase.get(i) + "\n");
            }

            writer.close();

        }catch (Exception e){
            logger.log(Level.WARNING, "Not Able To Save Case Details...: ", e);;
        }
    }

    private void customerWOHTable(TableView<CaseTableView> tableCustomers, boolean bool) {

        int caseCount = 0;

        tableCustomers.setVisible(true);

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCustomers.setContextMenu(menu);


            // Selecting and Copy the Case Number to Clipboard
            tableCustomers.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCustomers);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
                }
            });

        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void customerTable(String columnSelect, String filter, TableView<CaseTableView> tableCustomers, Boolean bool) {

        int caseCount = 0;

        tableCustomers.setVisible(true);

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCustomers.setContextMenu(menu);

            // Selecting and Copy the Case Number to Clipboard
            tableCustomers.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCustomers);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCustomers, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
                }
            });

        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void productViewCasesQueued(TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b, String sev) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;
            HSSFCell cellVal6;

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Severity")){
                    mycaseSevRefCell = i;
                }
                if (filterColName.equals("Support Theater")){
                     mycaseRegRefCell = i;
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
                            cellVal5 = filtersheet.getRow(i).getCell(mycaseSevRefCell);
                            String severity = cellVal5.getStringCellValue();
                            cellVal6 = filtersheet.getRow(i).getCell(mycaseRegRefCell);
                            String caseregion = cellVal6.getStringCellValue();

                            if (reg_product){
                                if (b) {
                                    if ((productName.equals(setProd.get(j)) && owner.startsWith("PS ")) && regChoice.getValue().equals(caseregion)) {

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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

                                        tableCases.getItems().addAll(observableList);
                                        caseCount++;
                                        if (tableCases.getItems().size() >= caseCount + 1) {
                                            tableCases.getItems().removeAll(observableList);
                                        }
                                    }
                                } else{
                                    if(sev.equals("All")) {
                                        if ((productName.equals(setProd.get(j)) && (owner.startsWith("TS ") || owner.startsWith("Tech-Ops"))) &&
                                                regChoice.getValue().equals(caseregion)) {

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
                                            int ribDays = 0;
                                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                                    localDate, array.get(9), array.get(10),
                                                    array.get(11), array.get(12), array.get(13),
                                                    array.get(14), array.get(15), array.get(16),
                                                    array.get(17)));

                                            tableCases.getItems().addAll(observableList);
                                            caseCount++;
                                            if (tableCases.getItems().size() >= caseCount + 1) {
                                                tableCases.getItems().removeAll(observableList);
                                            }
                                        }
                                    }
                                    if (sev.equals("BC")){
                                        if ((severity.equals("Business Critical") && (productName.equals(setProd.get(j)) && (owner.startsWith("PS") ||
                                                owner.startsWith("TS ") || owner.startsWith("Tech-Ops")))) && regChoice.getValue().equals(caseregion)){

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
                                            int ribDays = 0;
                                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                                    localDate, array.get(9), array.get(10),
                                                    array.get(11), array.get(12), array.get(13),
                                                    array.get(14), array.get(15), array.get(16),
                                                    array.get(17)));

                                            tableCases.getItems().addAll(observableList);
                                            caseCount++;
                                            if (tableCases.getItems().size() >= caseCount + 1) {
                                                tableCases.getItems().removeAll(observableList);
                                            }
                                        }
                                    }
                                    if (sev.equals("Major")){
                                        if ((severity.equals("Major") && (productName.equals(setProd.get(j)) && (owner.startsWith("PS") ||
                                                owner.startsWith("TS ") || owner.startsWith("Tech-Ops")))) && regChoice.getValue().equals(caseregion)){

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
                                            int ribDays = 0;
                                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                                    localDate, array.get(9), array.get(10),
                                                    array.get(11), array.get(12), array.get(13),
                                                    array.get(14), array.get(15), array.get(16),
                                                    array.get(17)));

                                            tableCases.getItems().addAll(observableList);
                                            caseCount++;
                                            if (tableCases.getItems().size() >= caseCount + 1) {
                                                tableCases.getItems().removeAll(observableList);
                                            }
                                        }
                                    }
                                    if (sev.equals("Minor")){
                                        if ((severity.equals("Minor") && (productName.equals(setProd.get(j)) && (owner.startsWith("PS") ||
                                                owner.startsWith("TS ") || owner.startsWith("Tech-Ops")))) && regChoice.getValue().equals(caseregion)){

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
                                            int ribDays = 0;
                                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                                    localDate, array.get(9), array.get(10),
                                                    array.get(11), array.get(12), array.get(13),
                                                    array.get(14), array.get(15), array.get(16),
                                                    array.get(17)));

                                            tableCases.getItems().addAll(observableList);
                                            caseCount++;
                                            if (tableCases.getItems().size() >= caseCount + 1) {
                                                tableCases.getItems().removeAll(observableList);
                                            }
                                        }
                                    }
                                }

                            }else{
                                if (b) {
                                    if ((productName.equals(setProd.get(j)) && owner.startsWith("PS ")) && regChoice.getValue().equals(caseregion)){

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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

                                        tableCases.getItems().addAll(observableList);
                                        caseCount++;
                                        if (tableCases.getItems().size() >= caseCount + 1) {
                                            tableCases.getItems().removeAll(observableList);
                                        }
                                    }
                                } else{
                                    if(sev.equals("All")) {
                                        if ((productName.equals(setProd.get(j)) && (owner.startsWith("TS ") || owner.startsWith("Tech-Ops"))) &&
                                                regChoice.getValue().equals(caseregion)){

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
                                            int ribDays = 0;
                                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                                    localDate, array.get(9), array.get(10),
                                                    array.get(11), array.get(12), array.get(13),
                                                    array.get(14), array.get(15), array.get(16),
                                                    array.get(17)));

                                            tableCases.getItems().addAll(observableList);
                                            caseCount++;
                                            if (tableCases.getItems().size() >= caseCount + 1) {
                                                tableCases.getItems().removeAll(observableList);
                                            }
                                        }
                                    }
                                    if (sev.equals("BC")){
                                        if ((severity.equals("Business Critical") && (productName.equals(setProd.get(j)) && (owner.startsWith("PS") ||
                                                owner.startsWith("TS ") || owner.startsWith("Tech-Ops")))) && regChoice.getValue().equals(caseregion)){

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
                                            int ribDays = 0;
                                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                                    localDate, array.get(9), array.get(10),
                                                    array.get(11), array.get(12), array.get(13),
                                                    array.get(14), array.get(15), array.get(16),
                                                    array.get(17)));

                                            tableCases.getItems().addAll(observableList);
                                            caseCount++;
                                            if (tableCases.getItems().size() >= caseCount + 1) {
                                                tableCases.getItems().removeAll(observableList);
                                            }
                                        }
                                    }
                                    if (sev.equals("Major")){
                                        if ((severity.equals("Major") && (productName.equals(setProd.get(j)) && (owner.startsWith("PS") ||
                                                owner.startsWith("TS ") || owner.startsWith("Tech-Ops")))) && regChoice.getValue().equals(caseregion)){

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
                                            int ribDays = 0;
                                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                                    localDate, array.get(9), array.get(10),
                                                    array.get(11), array.get(12), array.get(13),
                                                    array.get(14), array.get(15), array.get(16),
                                                    array.get(17)));

                                            tableCases.getItems().addAll(observableList);
                                            caseCount++;
                                            if (tableCases.getItems().size() >= caseCount + 1) {
                                                tableCases.getItems().removeAll(observableList);
                                            }
                                        }
                                    }
                                    if (sev.equals("Minor")){
                                        if ((severity.equals("Minor") && (productName.equals(setProd.get(j)) && (owner.startsWith("PS") ||
                                                owner.startsWith("TS ") || owner.startsWith("Tech-Ops")))) && regChoice.getValue().equals(caseregion)){

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
                                            int ribDays = 0;
                                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                    array.get(3), array.get(4), array.get(5), ribDays, age,
                                                    localDate, array.get(9), array.get(10),
                                                    array.get(11), array.get(12), array.get(13),
                                                    array.get(14), array.get(15), array.get(16),
                                                    array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }

    }

    private void productDueFilterView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, int dueDay, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;
            HSSFCell cellVal6;

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Support Theater)")) {
                    mycaseRegRefCell = i;
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
                            cellVal5 = filtersheet.getRow(i).getCell(rbnDaysRefCell);
                            String rbday = cellVal5.getStringCellValue();
                            int ribDays = Integer.parseInt(rbday);
                            cellVal6 = filtersheet.getRow(i).getCell(mycaseRegRefCell);
                            String caseregion = cellVal6.getStringCellValue();

                            if (b) {
                                if (((productName.equals(setProd.get(j)) && cellToCompare.equals(filter) && ribDays <= dueDay) &&
                                        (!caseStatus.equals("Pending Closure") || !caseStatus.equals("Future Availability"))) && regChoice.getValue().equals(caseregion)){

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

                                    int rbbnDays = 0;
                                    rbbnDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), rbbnDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if (((productName.equals(setProd.get(j)) && cellToCompare.equals(filter) && ribDays > dueDay)
                                        && (!caseStatus.equals("Pending Closure") || !caseStatus.equals("Future Availability"))) &&
                                        regChoice.getValue().equals(caseregion)){

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
                                    int rbbnDays = 0;
                                    rbbnDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), rbbnDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void inactiveCasesProductTableView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Support Theater")) {
                    mycaseRegRefCell = i;
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
                            cellVal4 = filtersheet.getRow(i).getCell(mycaseRegRefCell);
                            String caseregion = cellVal4.getStringCellValue();

                            if (reg_product){
                                if ((productName.equals(setProd.get(j)) && cellToCompare.equals(filter) &&
                                        (caseStatus.equals("Pending Closure") || caseStatus.equals("Future Availability"))) && regChoice.getValue().equals(caseregion)){

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                            else{
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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void twoFilterProductTableView(String columSelect1, String filter1, String columSelect2, String filter2, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Support Theater")) {
                    mycaseRegRefCell = i;
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
                            cellVal5 = filtersheet.getRow(i).getCell(mycaseRegRefCell);
                            String caseregion = cellVal5.getStringCellValue();

                            if (reg_product){
                                if ((productName.equals(setProd.get(j)) && cellToCompare.equals(filter1) && responsible.equals(filter2) &&
                                        (!caseStatus.equals("Pending Closure") || (!caseStatus.equals("Future Availability")))) && regChoice.getValue().equals(caseregion)){

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }

                            }else{
                                if (productName.equals(setProd.get(j)) && cellToCompare.equals(filter1) && responsible.equals(filter2) &&
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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }

    }

    private void productWIPCaseView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Support Theater")) {
                    mycaseRegRefCell = i;
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
                            cellVal4 = filtersheet.getRow(i).getCell(mycaseRegRefCell);
                            String caseregion = cellVal4.getStringCellValue();

                            if (reg_product){
                                if ((productName.equals(setProd.get(j)) && cellToCompare.equals(filter) && (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault")))) &&
                                        regChoice.getValue().equals(caseregion)){

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                            else{
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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void prodWOHTable(TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal3;
            HSSFCell cellVal4;

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Support Theater")) {
                    mycaseRegRefCell = i;
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
                            cellVal4 = filtersheet.getRow(i).getCell(mycaseRegRefCell);
                            String caseregion = cellVal4.getStringCellValue();

                            if (reg_product){
                                if (b) {
                                    if ((prodName.equals(setProd.get(j)) && (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) &&
                                            regChoice.getValue().equals(caseregion)){

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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

                                        tableCases.getItems().addAll(observableList);
                                        caseCount++;
                                        if (tableCases.getItems().size() >= caseCount + 1) {
                                            tableCases.getItems().removeAll(observableList);
                                        }
                                    }
                                } else {
                                    if ((prodName.equals(setProd.get(j)) && (caseStatus.equals("Pending Closure") || (caseStatus.equals("Future Availability")))) &&
                                            regChoice.getValue().equals(caseregion)) {

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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

                                        tableCases.getItems().addAll(observableList);
                                        caseCount++;
                                        if (tableCases.getItems().size() >= caseCount + 1) {
                                            tableCases.getItems().removeAll(observableList);
                                        }
                                    }
                                }
                            }
                            else{
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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }
                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void productOneFilterView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, Boolean bool) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Support Theater")) {
                    mycaseRegRefCell = i;
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
                            cellVal4 = filtersheet.getRow(i).getCell(mycaseRegRefCell);
                            String caseregion = cellVal4.getStringCellValue();

                            if (reg_product){
                                if (bool) {
                                    if (((product.equals(setProduct.get(j)) && cellValToCompare.equals(filter)) && (!caseStatus.equals("Pending Closure") &&
                                            (!caseStatus.equals("Future Availability")))) && regChoice.getValue().equals(caseregion)){

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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

                                        tableCases.getItems().addAll(observableList);
                                        caseCount++;
                                        if (tableCases.getItems().size() >= caseCount + 1) {
                                            tableCases.getItems().removeAll(observableList);
                                        }
                                    }
                                } else {
                                    if (((product.equals(setProduct.get(j)) && !cellValToCompare.equals(filter)) && (!caseStatus.equals("Pending Closure") &&
                                            (!caseStatus.equals("Future Availability")))) && regChoice.getValue().equals(caseregion)){

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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

                                        tableCases.getItems().addAll(observableList);
                                        caseCount++;
                                        if (tableCases.getItems().size() >= caseCount + 1) {
                                            tableCases.getItems().removeAll(observableList);
                                        }
                                    }
                                }
                            }
                            else{
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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }
                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void connectOkta() {

        if (!btnLogin.getText().equals("Logged!")) {

            WebEngine webEngine = webviewTest.getEngine();

            /*Client client = Clients.builder()
                    .setOrgUrl("https://dev-595242.oktapreview.com")
                    .setClientCredentials(new TokenClientCredentials("00G1rJSpdbBQ9xYwVFXtXfENnzTk4BaOoOIe8UdUma"))
                    .build();

            UserList users = client.listUsers();
            System.out.println(users);

            ApplicationList applications = client.listApplications();
            System.out.println(applications);

            Application app = client.getApplication("0oahhbtnnyXxx6rdE0h7");
            System.out.println(app);

            webviewTest.setContextMenuEnabled(true);
            com.sun.javafx.webkit.WebConsoleListener.setDefaultListener(
                    (webView, message, lineNumber, sourceId)-> System.out.println("Console: [" + sourceId + ":" + lineNumber + "] " + message)
            );*/

            webEngine.load("https://sonus.okta.com");
            logger.info("Connecting Sonus Okta...");

            browserLoginPane.toFront();
            apnBrowser.toFront();
            webEngine.load("https://sonus.okta.com");
            progressBar.setVisible(true);
            progressBar.toFront();
            progressBar.setProgress(0.20);
            lblDownload.setText(" CONNECTING/DOWNLOADING...");

            webEngine.getLoadWorker().stateProperty().addListener(new ChangeListener<Worker.State>() {
                @Override
                public void changed(ObservableValue ov, Worker.State oldState, Worker.State newState) {

                    if (newState == Worker.State.SUCCEEDED) {
                        if (webEngine.getLocation().equals("https://sonus.okta.com/app/UserHome")) {
                            logger.info("Login Success to Okta...");
                            progressBar.setProgress(0.40);
                            webEngine.load("https://sonus.okta.com/home/salesforce/0oayiqwes0HuzLJ6a1t6/46?fromHome=true");
                            progressBar.setProgress(0.50);
                            apnBrowser.toBack();
                            btnLogin.setDisable(true);
                            progressBar.setProgress(0.60);
                            webEngine.setJavaScriptEnabled(true);

                        }
                        if (webEngine.getLocation().contains("salesforce.com/500") || webEngine.getLocation().contains("salesforce.com/home")
                                || webEngine.getLocation().contains("lightning/page/home") || webEngine.getLocation().contains("rbbn.lightning.force.com/one/one.app")){

                                //webEngine.getLocation().equals("https://rbbn.my.salesforce.com/500/o") || webEngine.getLocation().equals("https://rbbn.my.salesforce.com/home/home.jsp") ||
                                //webEngine.getLocation().equals("https://na104.salesforce.com/500/o") || webEngine.getLocation().equals("https://na104.salesforce.com/home/home.jsp")) {

                            progressBar.setProgress(0.80);
                            btnLoadData.setVisible(true);
                            progressBar.setVisible(false);
                            btnLogin.setText("Logged!");
                            btnLogin.setVisible(false);
                            lblDownload.setText("CONNECTED - ONLINE");
                            logger.info("Connected to SalesForce, starting report download...");
                            lblDownload.setText("Downloading New Data!");
                            downData.downloadCSV();
                            readTimeStamp();
                            lblDownload.setText("CONNECTED - ONLINE!");
                            progressBar.setProgress(1);
                            apnMyCases.toFront();
                            myCasesPage();
                            // regionChoice();
                            lblStatus.setText("MY CASES");
                        }
                    }
                }
            });
        }
    }

    private void projectsPage() {

        HSSFCell theater;
        HSSFCell supHotReason;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_projects.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();

            prjAmericas = 0;
            prjEmea = 0;
            prjApac = 0;
            prjJapan = 0;
            prjGatingNow = 0;
            prjGatingDate = 0;
            prjGatingPrev = 0;
            prjAllCases = 0;

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                switch (filterColName) {
                    case ("Support Theater"):
                        caseRegionRef = i;
                        break;
                    case ("Support Hotlist Reason"):
                        caseSupHotListRRef = i;
                        break;
                }
            }

            for (int i = 1; i < lastRow + 1; i++) {

                theater = filtersheet.getRow(i).getCell(caseRegionRef);
                String region = theater.getStringCellValue();

                supHotReason = filtersheet.getRow(i).getCell(caseSupHotListRRef);
                String hotReason = supHotReason.getStringCellValue();

                if (region.equals("AMERICAS") || region.equals("NA")) {
                    prjAmericas++;
                }
                if (region.equals("EMEA")) {
                    prjEmea++;
                }
                if (region.equals("ASIAPAC")) {
                    prjApac++;
                }
                if (region.equals("JAPAN")) {
                    prjJapan++;
                }
                if (hotReason.equals("Project Gating - Now") || hotReason.equals("Gating Now")){
                    prjGatingNow++;
                }
                if (hotReason.equals("Project Gating - Date")){
                    prjGatingDate++;
                }
                if (hotReason.equals("Project Gating - Previously") || hotReason.equals("Previously Gating")){
                    prjGatingPrev++;
                }
                prjAllCases++;
            }
            btnAmericas.setText(String.valueOf(prjAmericas));
            btnEmea.setText(String.valueOf(prjEmea));
            btnApac.setText(String.valueOf(prjApac));
            btnJapan.setText(String.valueOf(prjJapan));
            btnGatingNow.setText(String.valueOf(prjGatingNow));
            btnGatingDate.setText(String.valueOf(prjGatingDate));
            btnGatingPrevious.setText(String.valueOf(prjGatingPrev));
            btnProjectsAll.setText(String.valueOf(prjAllCases));

        } catch (Exception e) {
            logger.log(Level.WARNING, "Project Page Build Failed...:", e);
        }
    }

    @FXML
    private void handleProjectClicks(MouseEvent event) throws IOException, InvalidFormatException{

        if (event.getSource() == btnAmericas){

            if (prjAmericas != 0) {
                initProjectTable();
                tableProjects.getItems().clear();
                tableProjects.setVisible(true);
                tableProjects.toFront();
                buildTableProjects(tableProjects, "NA");
            }
            if (prjAmericas == 0){
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnEmea){

            if (prjEmea != 0) {
                initProjectTable();
                tableProjects.getItems().clear();
                tableProjects.setVisible(true);
                tableProjects.toFront();
                buildTableProjects(tableProjects, "EMEA");
            }
            if (prjEmea == 0){
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnApac){

            if (prjApac != 0) {
                initProjectTable();
                tableProjects.getItems().clear();
                tableProjects.setVisible(true);
                tableProjects.toFront();
                buildTableProjects(tableProjects, "APAC");
            }
            if (prjApac == 0){
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnJapan){
            if (prjJapan != 0) {
                initProjectTable();
                tableProjects.getItems().clear();
                tableProjects.setVisible(true);
                tableProjects.toFront();
                buildTableProjects(tableProjects, "JAPAN");
            }
            if (prjJapan == 0){
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnGatingNow){
            if (prjGatingNow != 0) {
                initProjectTable();
                tableProjects.getItems().clear();
                tableProjects.setVisible(true);
                tableProjects.toFront();
                buildTableProjects(tableProjects, "NOW");
            }
            if (prjGatingNow == 0){
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnGatingPrevious){
            if (prjGatingPrev != 0) {
                initProjectTable();
                tableProjects.getItems().clear();
                tableProjects.setVisible(true);
                tableProjects.toFront();
                buildTableProjects(tableProjects, "PREV");
            }
            if (prjGatingPrev == 0){
                alertUser(strAlert);
            }
        }
        if (event.getSource() == btnGatingDate){
            if (prjGatingDate != 0) {
                initProjectTable();
                tableProjects.getItems().clear();
                tableProjects.setVisible(true);
                tableProjects.toFront();
                buildTableProjects(tableProjects, "DATE");
            }
            if (prjGatingDate == 0){
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnProjectsAll){
            if (prjAllCases !=0){
                tableProjects.getItems().clear();
                initProjectTable();
                buildTableProjects(tableProjects, "All");
            }
            if (prjAllCases == 0){
                alertUser(strAlert);
            }
        }

        if (event.getSource() == btnPrjMyNotes){
            getPrjNotesList();
        }
    }

    private void getPrjNotesList(){

        lstPrjNotes.getItems().clear();
        ArrayList<String> details = new ArrayList<String>();
        ObservableList<String> caseNotes = FXCollections.observableArrayList();

        File rep = new File(noteFolder + "\\Project");

        if (!rep.exists()){
            new File(noteFolder + "\\Project").mkdir();
        }else {
            File[] fileList = rep.listFiles();

            if (fileList.length == 0){
                String strNoNote = "THERE IS NO PERSONAL MEMO..." + "\n" + "\n" + "PLEASE CREATE PERSONAL MEMO FIRST!";
                alertUser(strNoNote);
            }else {

                for (int i = 0; i < fileList.length; i++) {
                    caseNotes.addAll(fileList[i].getName());
                }

                pnPrjNotes.setVisible(true);
                lstPrjNotes.getItems().addAll(caseNotes);
                lstPrjNotes.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);

                lstPrjNotes.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        txtPrjNoteView.clear();

                        String selected = lstPrjNotes.getSelectionModel().getSelectedItem().toString();
                        File prjCase = new File(noteFolder + "\\Project\\" + selected);

                        if (prjCase.isFile()) {
                            Scanner s = null;
                            try {
                                s = new Scanner(prjCase);
                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Projects - No File Found...: ", e);
                            }
                            while (s.hasNextLine()) {
                                txtPrjNoteView.appendText(s.nextLine() + "\n");
                            }
                            s.close();
                        }
                        btnPrjDelNote.setVisible(true);
                        btnPrjNewNote.setVisible(true);
                        txtPrjNoteView.setVisible(true);

                        btnPrjDelNote.setOnMouseClicked(new EventHandler<MouseEvent>() {
                            @Override
                            public void handle(MouseEvent event) {

                                File caseNoteFile = new File(noteFolder + "\\Project\\" + selected);
                                caseNoteFile.delete();

                                lstPrjNotes.getItems().remove(selected);

                                if (lstPrjNotes.getItems().size() == 0) {

                                    pnPrjNotes.setVisible(false);
                                    btnPrjNewNote.setVisible(false);
                                    btnPrjDelNote.setVisible(false);
                                    txtPrjNoteView.setVisible(false);
                                }
                                txtPrjNoteView.clear();
                            }
                        });

                        btnPrjNewNote.setOnMouseClicked(new EventHandler<MouseEvent>() {
                            @Override
                            public void handle(MouseEvent event) {
                                newProjectNote();
                            }
                        });
                    }
                });
            }
        }
    }

    private void newProjectNote(){

        Parent root;
        try {
            root = FXMLLoader.load(getClass().getClassLoader().getResource("home/CaseNoteProjects.fxml"));
            Stage stage = new Stage();
            stage.setTitle("ADD PERSONAL CASE NOTE");
            stage.getIcons().add(new Image("home/image/rbbicon.png"));
            stage.setScene(new Scene(root, 650, 400));
            stage.show();
            stage.setMinWidth(650);
            stage.setMinHeight(420);
            stage.setMaxWidth(650);
            stage.setMaxHeight(420);

        }
        catch (IOException e) {
            logger.log(Level.WARNING, "Unable to open new Project Note Window...: ", e);
        }
    }

    private void buildTableProjects(TableView<ProjectTableView> tableView, String str1){

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_projects.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell reason;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
                if (filterColName.equals("Support Hotlist Reason")){
                    caseSupHotListRRef = i;
                }
            }

            for (int k = 1; k < lastRow + 1; k++) {

                cellVal = filtersheet.getRow(k).getCell(caseRegionRef);
                String cellValToCompare = cellVal.getStringCellValue();
                reason = filtersheet.getRow(k).getCell(caseSupHotListRRef);
                String hotReason = reason.getStringCellValue();

                if (str1.equals("All")){

                    ArrayList<String> array = new ArrayList<>();
                    ObservableList<ProjectTableView> observableList = FXCollections.observableArrayList();

                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                    while (iterCells.hasNext()) {
                        HSSFCell cell = (HSSFCell) iterCells.next();
                        array.add(cell.getStringCellValue());
                    }


                    observableList.add(new ProjectTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), array.get(5), array.get(6), array.get(7),
                            array.get(8), array.get(9), array.get(10)));

                    tableView.getItems().addAll(observableList);
                    caseCount++;

                    if (tableView.getItems().size() >= caseCount + 1) {
                        tableView.getItems().removeAll(observableList);
                    }

                }

                if (str1.equals("NA")) {

                    if ((cellValToCompare.equals("AMERICAS") || cellValToCompare.equals("NA"))) {

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<ProjectTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        observableList.add(new ProjectTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), array.get(6), array.get(7),
                                array.get(8), array.get(9), array.get(10)));

                        tableView.getItems().addAll(observableList);
                        caseCount++;

                        if (tableView.getItems().size() >= caseCount + 1) {
                            tableView.getItems().removeAll(observableList);
                        }
                    }
                }
                if (str1.equals("EMEA")) {

                    if (cellValToCompare.equals("EMEA")) {

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<ProjectTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        observableList.add(new ProjectTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), array.get(6), array.get(7),
                                array.get(8), array.get(9), array.get(10)));

                        tableView.getItems().addAll(observableList);
                        caseCount++;

                        if (tableView.getItems().size() >= caseCount + 1) {
                            tableView.getItems().removeAll(observableList);
                        }
                    }
                }
                if (str1.equals("APAC")) {

                    if (cellValToCompare.equals("ASIAPAC")) {

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<ProjectTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        observableList.add(new ProjectTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), array.get(6), array.get(7),
                                array.get(8), array.get(9), array.get(10)));

                        tableView.getItems().addAll(observableList);
                        caseCount++;

                        if (tableView.getItems().size() >= caseCount + 1) {
                            tableView.getItems().removeAll(observableList);
                        }
                    }
                }
                if (str1.equals("JAPAN")){

                    if (cellValToCompare.equals("JAPAN")) {

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<ProjectTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        observableList.add(new ProjectTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), array.get(6), array.get(7),
                                array.get(8), array.get(9), array.get(10)));

                        tableView.getItems().addAll(observableList);
                        caseCount++;

                        if (tableView.getItems().size() >= caseCount + 1) {
                            tableView.getItems().removeAll(observableList);
                        }
                    }
                }
                if (str1.equals("NOW")) {

                    if ((hotReason.equals("Project Gating - Now") || hotReason.equals("Gating Now"))) {

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<ProjectTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        observableList.add(new ProjectTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), array.get(6), array.get(7),
                                array.get(8), array.get(9), array.get(10)));

                        tableView.getItems().addAll(observableList);
                        caseCount++;

                        if (tableView.getItems().size() >= caseCount + 1) {
                            tableView.getItems().removeAll(observableList);
                        }
                    }
                }
                if (str1.equals("PREV")) {

                    if ((hotReason.equals("Project Gating - Previously") || hotReason.equals("Previously Gating"))) {

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<ProjectTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }

                        observableList.add(new ProjectTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), array.get(6), array.get(7),
                                array.get(8), array.get(9), array.get(10)));

                        tableView.getItems().addAll(observableList);
                        caseCount++;

                        if (tableView.getItems().size() >= caseCount + 1) {
                            tableView.getItems().removeAll(observableList);
                        }
                    }
                }
                if (str1.equals("DATE")) {

                    if (hotReason.equals("Project Gating - Date")) {

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<ProjectTableView> observableList = FXCollections.observableArrayList();

                        Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                        while (iterCells.hasNext()) {
                            HSSFCell cell = (HSSFCell) iterCells.next();
                            array.add(cell.getStringCellValue());
                        }


                        observableList.add(new ProjectTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), array.get(6), array.get(7),
                                array.get(8), array.get(9), array.get(10)));

                        tableView.getItems().addAll(observableList);
                        caseCount++;

                        if (tableView.getItems().size() >= caseCount + 1) {
                            tableView.getItems().removeAll(observableList);
                        }
                    }
                }
            }

            btnToExcel.setVisible(true);
            apnProjects.toFront();
            btnToExcel.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    exportExcelActionProjects(tableView);
                }
            });

            menu = new ContextMenu();
            String caseno = "";
            menu.getItems().add(openCaseSFDC);
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseDetails);
            tableView.setContextMenu(menu);

            // Selecting and Copy the Case Number to Clipboard
            tableView.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboardProjects(tableView);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    projectCaseDetails();
                }
            });

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newProjectNote();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumberProjects(tableView, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }
                }
            });


        } catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void projectCaseDetails(){
        Parent root;

        try {
            root = FXMLLoader.load(getClass().getClassLoader().getResource("home/CaseDetails.fxml"));
            Stage stage = new Stage();
            stage.setTitle("CASE DETAILS WINDOW");
            stage.getIcons().add(new Image("home/image/rbbicon.png"));

            if (screenHeight > 1025) {
                stage.setScene(new Scene(root, 640, 940));
                stage.show();
                stage.setMinWidth(640);
                stage.setMinHeight(940);
                stage.setMaxWidth(640);
                stage.setMaxHeight(940);
            }
            if (screenHeight <1025){
                stage.setScene(new Scene(root, 640, 660));
                stage.show();
                stage.setMinWidth(640);
                stage.setMinHeight(660);
                stage.setMaxWidth(640);
                stage.setMaxHeight(660);
            }
        }
        catch (IOException e) {
            logger.log(Level.WARNING, "Unable to open Project Case Details Window...", e);
        }
    }

    private void myCaseDetails(){

        Parent root;

        try {

            root = FXMLLoader.load(getClass().getClassLoader().getResource("home/MyCaseDetails.fxml"));
            Stage stage = new Stage();
            stage.setTitle("CASE DETAILS WINDOW");
            stage.getIcons().add(new Image("home/image/rbbicon.png"));

            if (screenHeight > 1025) {
                stage.setScene(new Scene(root, 740, 920));
                stage.show();
                stage.setMinWidth(740);
                stage.setMinHeight(920);
                stage.setMaxWidth(740);
                stage.setMaxHeight(920);
            }
            if (screenHeight <1025){
                stage.setScene(new Scene(root, 740, 660));
                stage.show();
                stage.setMinWidth(740);
                stage.setMinHeight(660);
                stage.setMaxWidth(740);
                stage.setMaxHeight(660);
            }

        }
        catch (IOException e) {
            logger.log(Level.WARNING, "Unable to open Case Details Window...", e);
        }

        //saveCaseDetails();
    }


    private void writeDefaultSettingsToFile(String userFilter, String queueFilter, String productFilter, String regionFilter, String accountFilter, String wgFilter) {

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
        Collections.sort(settingsUsers);

        try {

            FileWriter writer = new FileWriter(new File(settingsFolder + "\\cmt_user_default_settings.txt"));
            int size = settingsUsers.size();
            for (int i = 0; i < size; i++) {
                String str = settingsUsers.get(i);
                writer.write(str);
                if (i < size - 1)
                    writer.write("\n");
            }

            writer.close();

        } catch (Exception e) {
            logger.log(Level.WARNING, "User Default Settings Save Failed!", e);
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
        Collections.sort(settingsQueue);

        try {

            FileWriter writer = new FileWriter(new File(settingsFolder + "\\cmt_queueu_default_settings.txt"));
            int size = settingsQueue.size();
            for (int i = 0; i < size; i++) {
                String str = settingsQueue.get(i);
                writer.write(str);
                if (i < size - 1)
                    writer.write("\n");
            }
            writer.close();

        } catch (Exception e) {
            logger.log(Level.WARNING, "Queue Default Settings Save Failed!", e);
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
        Collections.sort(settingsProducts);

        try {

            FileWriter writer = new FileWriter(new File(settingsFolder + "\\cmt_product_default_settings.txt"));
            int size = settingsProducts.size();
            for (int i = 0; i < size; i++) {
                String str = settingsProducts.get(i);
                writer.write(str);
                if (i < size - 1)
                    writer.write("\n");
            }
            writer.close();

        } catch (Exception e) {
            logger.log(Level.WARNING, "Products Default Settings Save Failed!", e);
        }

        ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
        int accArraySize = setAcc.size();

        settingsAccounts = (ArrayList<String>) setAcc.stream().distinct().collect(Collectors.toList());
        Collections.sort(settingsAccounts);

        try {

            FileWriter writer = new FileWriter(new File(settingsFolder + "\\cmt_account_default_settings.txt"));
            int size = settingsAccounts.size();
            for (int i = 0; i < size; i++) {
                String str = settingsAccounts.get(i);
                writer.write(str);
                if (i < size - 1)
                    writer.write("\n");
            }
            writer.close();

        } catch (Exception e) {
            logger.log(Level.WARNING, "Account Default Settings Save Failed!", e);
        }

        try {

            FileWriter writer = new FileWriter(new File(settingsFolder + "\\cmt_region_default_settings.txt"));
            String str = regChoice.getValue().toString();
            String check = Boolean.toString(reg_product);
            writer.write(str);
            writer.write("\n");
            writer.write(check);
            writer.close();

        } catch (Exception e) {
            logger.log(Level.WARNING, "Region Default Settings Save Failed!", e);
        }

        ArrayList<String> setWg = new ArrayList<>(Arrays.asList(txWorkGroup.getText().split(",\\s*")));
        settingsWorkGroups = (ArrayList<String>) setWg.stream().distinct().collect(Collectors.toList());
        Collections.sort(settingsWorkGroups);

        try {

            FileWriter writer = new FileWriter(new File(settingsFolder + "\\cmt_wg_default_settings.txt"));
            int size = settingsWorkGroups.size();
            for (int i = 0; i < size; i++) {
                String str = settingsWorkGroups.get(i);
                writer.write(str);
                if (i < size - 1)
                    writer.write("\n");
            }
            writer.close();

        } catch (Exception e) {
            logger.log(Level.WARNING, "WG Default Settings Save Failed!", e);
        }
    }

    private void readTimeStamp(){

        File timeStampFile = new File(dataFolder + "\\cmt_data_Date.txt");

        if (timeStampFile.isFile()){
            Scanner s = null;
            try{

                s = new Scanner(timeStampFile);

            }catch (Exception e){
                logger.log(Level.WARNING, "Read Time Stamp File Failed!", e);
            }

            ArrayList<String> readDate = new ArrayList<>();
            while(s.hasNextLine()){
                readDate.add(s.nextLine());
            }
            s.close();

            lblRefreshText.setVisible(true);
            lblRefreshText.setText(readDate.get(0)+ "\n" + readDate.get(1) + "\n" + readDate.get(2));
        }

        timeData = new Timeline();
        timeData.setCycleCount(Timeline.INDEFINITE);
        timeData.getKeyFrames().add(new KeyFrame(Duration.minutes(6), new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                timeData.stop();
                logger.info("Refreshing Time-Stamps!...");
                readTimeStamp();
            }
        }));
        timeData.playFromStart();
    }

    private void readDefaultSettingFiles() {

        // Load Already Saved Settings File if there are any

        File settingUsersFile = new File(settingsFolder + "\\cmt_user_default_settings.txt");
        File settingQueueFile = new File(settingsFolder + "\\cmt_queueu_default_settings.txt");
        File settingProductsFile = new File(settingsFolder + "\\cmt_product_default_settings.txt");
        File settingAccountsFile = new File(settingsFolder + "\\cmt_account_default_settings.txt");
        File settingRegionFile = new File(settingsFolder + "\\cmt_region_default_settings.txt");
        File settingWGFile = new File(settingsFolder + "\\cmt_wg_default_settings.txt");

        if (settingUsersFile.isFile()) {

            Scanner s = null;
            try {
                s = new Scanner(settingUsersFile);
            } catch (FileNotFoundException e) {
                logger.log(Level.WARNING, "No Saved User List", e);
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
                logger.log(Level.WARNING, "No Saved Queue List", e);
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
                logger.log(Level.WARNING, "No Saved Product List", e);
            }
            ArrayList<String> readProductList = new ArrayList<String>();
            while (s.hasNextLine()) {
                readProductList.add(s.nextLine());
            }
            s.close();

            txProducts.setText(readProductList.stream().collect(Collectors.joining(", ")));
        }

        if (settingAccountsFile.isFile()) {

            Scanner s = null;
            try {
                s = new Scanner(settingAccountsFile);
            } catch (FileNotFoundException e) {
                logger.log(Level.WARNING, "No Saved Account List", e);
            }
            ArrayList<String> readAccountList = new ArrayList<String>();
            while (s.hasNextLine()) {
                readAccountList.add(s.nextLine());
            }
            s.close();

            txAccounts.setText(readAccountList.stream().collect(Collectors.joining(", ")));
        }

        if (settingRegionFile.isFile()) {

            Scanner s = null;
            try {
                s = new Scanner(settingRegionFile);
            } catch (FileNotFoundException e) {
                logger.log(Level.WARNING, "No Saved Region List", e);
            }
            ArrayList<String> readRegion = new ArrayList<String>();
            while (s.hasNextLine()) {
                readRegion.add(s.nextLine());
            }
            s.close();
            if (readRegion.size() > 1) {
                regChoice.setValue(readRegion.get(0));
                reg_product = Boolean.parseBoolean(readRegion.get(1));
                if (reg_product){
                    checkRegProduct.setSelected(true);
                }

            }
        }

        if (settingWGFile.isFile()) {

            Scanner s = null;
            try {
                s = new Scanner(settingWGFile);
            } catch (FileNotFoundException e) {
                logger.log(Level.WARNING, "No Saved Account List", e);
            }
            ArrayList<String> readWgList = new ArrayList<String>();
            while (s.hasNextLine()) {
                readWgList.add(s.nextLine());
            }
            s.close();

            txWorkGroup.setText(readWgList.stream().collect(Collectors.joining(", ")));
        }
    }

    private void caseUpdateTableView(String caseTableSelect, TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b, boolean bool) {

        int caseCount = 0;

        LocalDate dateToday = LocalDate.now();
        LocalDate caseUpdateDate = null;
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");


        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays,age,
                                    caseUpdateDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

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
                            int ribDays = 0;
                            ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                            observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                    array.get(3), array.get(4), array.get(5), ribDays,age,
                                    caseUpdateDate, array.get(9), array.get(10),
                                    array.get(11), array.get(12), array.get(13),
                                    array.get(14), array.get(15), array.get(16),
                                    array.get(17)));

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
                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                caseUpdateDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            // Select and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }

                }
            });

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void mycaseUpdateTableView(String caseTableSelect, TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b, boolean bool) {

        int caseCount = 0;

        LocalDate dateToday = LocalDate.now();
        LocalDate caseUpdateDate = null;
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");


        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellValStat;
            HSSFCell cellValUser;
            HSSFCell cellVal2;


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
                if (filterColName.equals("Co-Owner")){
                    myCoOwnCaseRefCell = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty()) || (!txWorkGroup.getText().isEmpty())) {

                setUsers = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));

                if(!txtWGNames.getText().isEmpty()) {
                    setWgNames.removeAll(setUsers);
                    setUsers.addAll(setWgNames);
                }

                int userfiltnum = setUsers.size();

                for (int i = 0; i < userfiltnum; i++) {

                    for (int k = 1; k < lastRow + 1; k++) {

                        cellVal = filtersheet.getRow(k).getCell(caseNextUpdateDateRef);
                        String cellValToCompare = cellVal.getStringCellValue();

                        cellValStat = filtersheet.getRow(k).getCell(caseStatRefCell);
                        String cellStat = cellValStat.getStringCellValue();

                        cellValUser = filtersheet.getRow(k).getCell(mycaseOwnerRefCell);
                        String caseUser = cellValUser.getStringCellValue();

                        cellVal2 = filtersheet.getRow(k).getCell(myCoOwnCaseRefCell);
                        String coOwner = cellVal2.getStringCellValue();

                        ArrayList<String> array = new ArrayList<>();
                        ObservableList<CaseTableView> observableList = FXCollections.observableArrayList();

                        if (!cellValToCompare.equals("NotSet")) {

                            caseUpdateDate = LocalDate.parse(cellValToCompare, formatter);
                        }

                        if ((b) && (bool)) {

                            if (((caseUser.equals(setUsers.get(i)) || coOwner.equals(setUsers.get(i)))&& !cellValToCompare.equals("NotSet"))) {

                                if ( caseUpdateDate.compareTo(dateToday) == 0) {

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(caseAgeRefCell));
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            caseUpdateDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
                        if ((!b) && (bool)) {

                            if (((caseUser.equals(setUsers.get(i)) || coOwner.equals(setUsers.get(i)))&& !cellValToCompare.equals("NotSet"))) {

                                if (caseUpdateDate.compareTo(dateToday) < 0) {

                                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                                    while (iterCells.hasNext()) {
                                        HSSFCell cell = (HSSFCell) iterCells.next();
                                        array.add(cell.getStringCellValue());
                                    }

                                    int age = 0;
                                    age = Integer.parseInt(array.get(caseAgeRefCell));
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            caseUpdateDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
                        if (!b && !bool) {

                            if (((caseUser.equals(setUsers.get(i)) || coOwner.equals(setUsers.get(i))) && cellValToCompare.equals("NotSet"))) {

                                Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                                while (iterCells.hasNext()) {
                                    HSSFCell cell = (HSSFCell) iterCells.next();
                                    array.add(cell.getStringCellValue());
                                }

                                caseUpdateDate = null;

                                int age = 0;
                                age = Integer.parseInt(array.get(caseAgeRefCell));
                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays,age,
                                        caseUpdateDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
                }
            });

            // Select and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void createMyQueueCaseView(String columnSelect, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays,age,
                                        localDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void oneFilterMyTableView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Co-Owwer")){
                    myCoOwnCaseRefCell = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty() || !txWorkGroup.getText().isEmpty())) {

                setUsers = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                //ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                //ArrayList<String> mergedOwner = new ArrayList<>();

                if (!txWorkGroup.getText().isEmpty()) {
                    setWgNames.removeAll(setUsers);
                    setUsers.addAll(setWgNames);
                }

                int userfiltnum = setUsers.size();
                //int userqueuenum = setQueu.size();

                /*if (!setUser.isEmpty()) {
                    for (int i = 0; i < userfiltnum; i++) {

                        mergedOwner.add(setUser.get(i));
                    }
                }

                if (!setQueu.isEmpty()) {
                    for (int i = 0; i < userqueuenum; i++) {
                        mergedOwner.add(setQueu.get(i));
                    }
                }*/

                //int mergedUserNum = mergedOwner.size();


                if ((!setUsers.isEmpty())) {

                    for (int j = 0; j < userfiltnum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(myCaseCellRef1);
                            String cellValToCompare = cellVal2.getStringCellValue();
                            cellVal4 = filtersheet.getRow(i).getCell(myCoOwnCaseRefCell);
                            String coOwner = cellVal4.getStringCellValue();

                            if (b) {
                                if (((caseUser.equals(setUsers.get(j)) || (coOwner.equals(setUsers.get(j)))) && cellValToCompare.equals(filter)) && (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) {

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if (((caseUser.equals(setUsers.get(j)) || coOwner.equals(setUsers.get(j)))&& !cellValToCompare.equals(filter)) && (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) {

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();
                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }
                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
                }
            });
            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void myWOHTableView(TableView<CaseTableView> tableCases, AnchorPane apnTableView, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Co-Owner")){
                    myCoOwnCaseRefCell = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty() || !txWorkGroup.getText().isEmpty())) {

                setUsers = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                //ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                //ArrayList<String> mergedOwner = new ArrayList<>();

                if (!txtWGNames.getText().isEmpty()) {
                    setWgNames.removeAll(setUsers);
                    setUsers.addAll(setWgNames);
                }

                int userfiltnum = setUsers.size();
                //int userqueuenum = setQueu.size();

                /*if (!setUser.isEmpty()) {
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
                */

                if ((!setUsers.isEmpty())) {

                    for (int j = 0; j < userfiltnum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(myCoOwnCaseRefCell);
                            String coOwner = cellVal2.getStringCellValue();

                            if (b) {
                                if ((caseUser.equals(setUsers.get(j)) || coOwner.equals(setUsers.get(j))) && (!caseStatus.equals("Pending Closure") && (!caseStatus.equals("Future Availability")))) {

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableCases.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableCases.getItems().size() >= caseCount + 1) {
                                        tableCases.getItems().removeAll(observableList);
                                    }
                                }
                            } else {
                                if ((caseUser.equals(setUsers.get(j)) || coOwner.equals(setUsers.get(j))) && (caseStatus.equals("Pending Closure") || (caseStatus.equals("Future Availability")))) {

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays,age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
                }
            });

            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void viewComments(String caseNum){

        ArrayList<String> caseCommentArray = new ArrayList();

        try(HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_comments.xls")))){

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int mycaseNumCellRef = 0;
            int myCaseCommentDateRef = 0;
            int myCaseCommentRef = 0;
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();

            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;

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

                if (commentCaseNumber.equals(caseNum)){

                    caseCommentArray.add(commentDate);
                    caseCommentArray.add(commentComment);
                }
            }

            int arraySize = caseCommentArray.size();

            if (arraySize == 0){
                Alert alert = new Alert(Alert.AlertType.WARNING);
                ((Stage) alert.getDialogPane().getScene().getWindow()).getIcons().add(new Image("home/image/rbbicon.png"));
                alert.setTitle("RBBN CMT WARNING:");
                alert.setHeaderText(null);
                alert.setContentText("THERE IS NO COMMENT FOR THIS CASE"+ "\n" + "SINCE 7 DAYS!");
                alert.showAndWait();
            }

            spnNote.setVisible(true);
            txtShowCaseNotes.clear();

            for (int i = 0; i < arraySize; i += 2) {

                txtShowCaseNotes.appendText("===============" + "\n" + caseCommentArray.get(i)+ "\n" + "\n" + caseCommentArray.get(i+1) + "\n");
            }
            txtShowCaseNotes.positionCaret(0);

        }catch (Exception e){
            logger.log(Level.WARNING, "Work Notes Build Failed!", e);
        }
        caseCommentArray.clear();
    }

    private void viewCaseComments() {
        Parent root;

        try {
            root = FXMLLoader.load(getClass().getClassLoader().getResource("home/CaseComment.fxml"));
            Stage stage = new Stage();
            stage.setTitle("VIEW CASE COMMENTS FROM LAST 7 DAYS");
            stage.getIcons().add(new Image("home/image/rbbicon.png"));
            stage.setScene(new Scene(root, 650, 400));
            stage.show();
            stage.setMinWidth(650);
            stage.setMinHeight(420);
            stage.setMaxWidth(650);
            stage.setMaxHeight(420);

            //saveCaseDetails();

        }
        catch (IOException e) {
            logger.log(Level.WARNING, "View Work Notes Window Failed...", e);
        }
    }

    private void overviewMyWIPCaseTableView(String columFilter, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Co-Owner")){
                    myCoOwnCaseRefCell = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty() || !txWorkGroup.getText().isEmpty())) {

                setUsers = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                //ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                //ArrayList<String> mergedOwner = new ArrayList<>();

                if(!txtWGNames.getText().isEmpty()) {
                    setWgNames.removeAll(setUsers);
                    setUsers.addAll(setWgNames);
                }

                int userfiltnum = setUsers.size();
                //int userqueuenum = setQueu.size();

                /*if (!setUser.isEmpty()) {
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
                   */

                if ((!setUsers.isEmpty())) {

                    for (int j = 0; j < userfiltnum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal4 = filtersheet.getRow(i).getCell(myCoOwnCaseRefCell);
                            String coOwner = cellVal4.getStringCellValue();

                            if ((caseUser.equals(setUsers.get(j)) || coOwner.equals(setUsers.get(j))) && cellToCompare.equals(filter) && (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault")))) {

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
                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays,age,
                                        localDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
                }
            });
            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void twoFilterMyTableView(String columSelect1, String filter1, String columSelect2, String filter2, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;

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
                if (filterColName.equals("Co-Owner")){
                    myCoOwnCaseRefCell = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty() || !txWorkGroup.getText().isEmpty())) {

                setUsers = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                //ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                //ArrayList<String> mergedOwner = new ArrayList<>();

                if (!txtWGNames.getText().isEmpty()) {
                    setWgNames.removeAll(setUsers);
                    setUsers.addAll(setWgNames);
                }

                int userfiltnum = setUsers.size();
                //int userqueuenum = setQueu.size();

                /*if (!setUser.isEmpty()) {
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
                */


                if ((!setUsers.isEmpty())) {

                    for (int j = 0; j < userfiltnum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal4 = filtersheet.getRow(i).getCell(caseCellRef2);
                            String responsible = cellVal4.getStringCellValue();
                            cellVal5 = filtersheet.getRow(i).getCell(myCoOwnCaseRefCell);
                            String coOwner = cellVal5.getStringCellValue();


                            if ((caseUser.equals(setUsers.get(j)) || coOwner.equals(setUsers.get(j))) && cellToCompare.equals(filter1) && responsible.equals(filter2) &&
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
                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays,age,
                                        localDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }
                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
                }
            });
            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void inactiveCasesMyTableView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty() || !txWorkGroup.getText().isEmpty())) {

                setUsers = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                //ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                //ArrayList<String> mergedOwner = new ArrayList<>();

                if(!txtWGNames.getText().isEmpty()) {
                    setWgNames.removeAll(setUsers);
                    setUsers.addAll(setWgNames);
                }

                int userfiltnum = setUsers.size();
                //int userqueuenum = setQueu.size();

                /*if (!setUser.isEmpty()) {
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
                */

                if ((!setUsers.isEmpty())) {

                    for (int j = 0; j < userfiltnum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal4 = filtersheet.getRow(i).getCell(myCoOwnCaseRefCell);
                            String coOwner = cellVal4.getStringCellValue();

                            if ((caseUser.equals(setUsers.get(j)) || coOwner.equals(setUsers.get(j))) && cellToCompare.equals(filter) &&
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
                                int ribDays = 0;
                                ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                        array.get(3), array.get(4), array.get(5), ribDays,age,
                                        localDate, array.get(9), array.get(10),
                                        array.get(11), array.get(12), array.get(13),
                                        array.get(14), array.get(15), array.get(16),
                                        array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
                }
            });
            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void myDueFilterView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, int dueDay, boolean b) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal1;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;
            HSSFCell cellVal5;
            HSSFCell cellVal6;


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
                if (filterColName.equals("Co-Owner")){
                    myCoOwnCaseRefCell = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")){
                    myRbbnDaysRefCell = i;
                }
            }

            if ((!txUsers.getText().isEmpty() || !txWorkGroup.getText().isEmpty())) {

                setUsers = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));
                //ArrayList<String> setQueu = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                //ArrayList<String> mergedOwner = new ArrayList<>();

                if(!txtWGNames.getText().isEmpty()) {
                    setWgNames.removeAll(setUsers);
                    setUsers.addAll(setWgNames);
                }

                int userfiltnum = setUsers.size();

                if ((!setUsers.isEmpty())) {

                    for (int j = 0; j < userfiltnum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            cellVal1 = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseUser = cellVal1.getStringCellValue();
                            cellVal2 = filtersheet.getRow(i).getCell(caseCellRef);
                            String cellToCompare = cellVal2.getStringCellValue();
                            cellVal3 = filtersheet.getRow(i).getCell(caseStatRefCell);
                            String caseStatus = cellVal3.getStringCellValue();
                            cellVal4 = filtersheet.getRow(i).getCell(mycaseAgeRefCell);
                            int compAge = Integer.parseInt(cellVal4.getStringCellValue());
                            cellVal5 = filtersheet.getRow(i).getCell(myCoOwnCaseRefCell);
                            String coOwner = cellVal5.getStringCellValue();
                            cellVal6 = filtersheet.getRow(i).getCell(myRbbnDaysRefCell);
                            String rbbnDays = cellVal6.getStringCellValue();
                            int myRbbnDays = Integer.parseInt(rbbnDays);


                            if ((cellToCompare.equals(filter))) {

                                if (b) {

                                    if (((caseUser.equals(setUsers.get(j))  || coOwner.equals(setUsers.get(j))) && myRbbnDays <= dueDay) && ((caseStatus.equals("Open / Assign") ||
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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

                                        tableCases.getItems().addAll(observableList);
                                        caseCount++;
                                        if (tableCases.getItems().size() >= caseCount + 1) {
                                            tableCases.getItems().removeAll(observableList);
                                        }
                                    }
                                } else {
                                    if (((caseUser.equals(setUsers.get(j))  || coOwner.equals(setUsers.get(j))) && myRbbnDays > dueDay) && ((caseStatus.equals("Open / Assign") ||
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
                                        int ribDays = 0;
                                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                                localDate, array.get(9), array.get(10),
                                                array.get(11), array.get(12), array.get(13),
                                                array.get(14), array.get(15), array.get(16),
                                                array.get(17)));

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
            menu.getItems().add(casePersonalNote);
            menu.getItems().add(openCaseComments);
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            casePersonalNote.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    newCaseNote();

                }});

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
                    }

                }
            });

            openCaseComments.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    viewCaseComments();
                }
            });
            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void oneFilterTableView(String columnSelect, String filter1, TableView tableCases, AnchorPane apnTableView, Boolean bool) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

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
                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

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
                menu.getItems().add(openCaseDetails);
                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void twoFilterTableView(String columnSelect1, String columnSelect2, String filter1, String filter2, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                    int ribDays = 0;
                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));
                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), array.get(5), ribDays,age,
                            localDate, array.get(9), array.get(10),
                            array.get(11), array.get(12), array.get(13),
                            array.get(14), array.get(15), array.get(16),
                            array.get(17)));

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
                menu.getItems().add(openCaseDetails);
                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void overviewWIPCaseTableView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                    int ribDays = 0;
                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), array.get(5), ribDays,age,
                            localDate, array.get(9), array.get(10),
                            array.get(11), array.get(12), array.get(13),
                            array.get(14), array.get(15), array.get(16),
                            array.get(17)));

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
                menu.getItems().add(openCaseDetails);

                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void overviewEngineeringTableView(String columnSelect, String filter1, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                    int ribDays = 0;
                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), array.get(5), ribDays,age,
                            localDate, array.get(9), array.get(10),
                            array.get(11), array.get(12), array.get(13),
                            array.get(14), array.get(15), array.get(16),
                            array.get(17)));

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
                menu.getItems().add(openCaseDetails);
                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void overViewInactiveTable(String columnSelect1, String filter1, TableView<CaseTableView> tableCases, AnchorPane apnTableView) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                    int ribDays = 0;
                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), array.get(5), ribDays,age,
                            localDate, array.get(9), array.get(10),
                            array.get(11), array.get(12), array.get(13),
                            array.get(14), array.get(15), array.get(16),
                            array.get(17)));

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
                menu.getItems().add(openCaseDetails);
                tableCases.setContextMenu(menu);

                openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        myCaseDetails();
                    }
                });

                openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                    @Override
                    public void handle(ActionEvent event) {
                        try {

                            String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                            URL caseSearch = new URL(search);
                            Desktop.getDesktop().browse(caseSearch.toURI());
                        }catch (Exception e){
                            logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                            logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void overviewWOHView(Boolean bool) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

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
                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void overviewSevQueueView(String columnSelect, String filter, TableView tableView, AnchorPane anchorpane){

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                    caseCellRef2 = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Case Owner")){
                    caseOwnerRefCell = i;
                }

            }
            for (int k = 1; k < lastRow + 1; k++) {

                cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                String cellValToCompare = cellVal.getStringCellValue();

                cellVal2 = filtersheet.getRow(k).getCell(caseOwnerRefCell);
                String owner = cellVal2.getStringCellValue();

                if (owner.startsWith("PS") || owner.startsWith("TS") || owner.startsWith("Tech-Ops")) {

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
                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays, age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

                        tableView.getItems().addAll(observableList);
                        caseCount++;
                        if (tableView.getItems().size() >= caseCount + 1) {
                            tableView.getItems().removeAll(observableList);
                        }
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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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



        }catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }
    private void mySevQueueView(String columnSelect, String filter, TableView tableView, AnchorPane anchorpane){

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                    caseCellRef2 = i;
                }
                if (filterColName.equals("Next Case Update")) {
                    caseNextUpdateDateRef = i;
                }
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
                }
                if (filterColName.equals("Case Owner")){
                    caseOwnerRefCell = i;
                }
            }

            if(!txQueues.getText().isEmpty()){

                String queueFilter = txQueues.getText();
                ArrayList<String> setQueue = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                int queuefiltnum = setQueue.size();

                if ((!queueFilter.equals(""))) {

                    for (int j = 0; j < queuefiltnum; j++) {

                        for (int k = 1; k < lastRow + 1; k++) {

                            cellVal = filtersheet.getRow(k).getCell(caseCellRef);
                            String cellValToCompare = cellVal.getStringCellValue();

                            cellVal2 = filtersheet.getRow(k).getCell(caseOwnerRefCell);
                            String owner = cellVal2.getStringCellValue();

                            if (owner.equals(setQueue.get(j))){

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
                                    int ribDays = 0;
                                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                            array.get(3), array.get(4), array.get(5), ribDays, age,
                                            localDate, array.get(9), array.get(10),
                                            array.get(11), array.get(12), array.get(13),
                                            array.get(14), array.get(15), array.get(16),
                                            array.get(17)));

                                    tableView.getItems().addAll(observableList);
                                    caseCount++;
                                    if (tableView.getItems().size() >= caseCount + 1) {
                                        tableView.getItems().removeAll(observableList);
                                    }
                                }
                            }
                        }
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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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



        }catch (Exception e) {
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void overviewQueueView(String columnSelect, String filter, TableView tableView, AnchorPane anchorpane, String overText) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbnDaysRefCell = i;
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
                    int ribDays = 0;
                    ribDays = Integer.parseInt(array.get(rbnDaysRefCell));

                    observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                            array.get(3), array.get(4), array.get(5), ribDays,age,
                            localDate, array.get(9), array.get(10),
                            array.get(11), array.get(12), array.get(13),
                            array.get(14), array.get(15), array.get(16),
                            array.get(17)));

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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
        }
    }

    private void overviewDueFilterView(String columnSelect, String filter, TableView<CaseTableView> tableCases, AnchorPane apnTableView, int ageDue, Boolean due) {

        int caseCount = 0;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();
            HSSFCell cellVal;
            HSSFCell cellVal2;
            HSSFCell cellVal3;
            HSSFCell cellVal4;

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
                if (filterColName.equals("Days by Ribbon (days)")) {
                    rbbnDaysRefCell = i;
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
                cellVal4 = filtersheet.getRow(k).getCell(rbbnDaysRefCell);
                String rbnDays = cellVal4.getStringCellValue();
                int rbbnAge = Integer.parseInt(rbnDays);

                if (due) {

                    if ((cellValToCompare.equals(filter) && rbbnAge <= ageDue) && ((!caseStatus.equals("Develop Solution") &&
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

                        int age;
                        age = Integer.parseInt(array.get(caseAgeRefCell));
                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

                        tableCases.getItems().addAll(observableList);

                        caseCount++;

                        if (tableCases.getItems().size() >= caseCount + 1) {
                            tableCases.getItems().removeAll(observableList);
                        }
                    }
                } else {
                    if ((cellValToCompare.equals(filter) && rbbnAge > ageDue) && ((!caseStatus.equals("Develop Solution") &&
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

                        int age;
                        age = Integer.parseInt(array.get(caseAgeRefCell));
                        int ribDays = 0;
                        ribDays = Integer.parseInt(array.get(rbbnDaysRefCell));

                        observableList.add(new CaseTableView(array.get(0), array.get(1), array.get(2),
                                array.get(3), array.get(4), array.get(5), ribDays,age,
                                localDate, array.get(9), array.get(10),
                                array.get(11), array.get(12), array.get(13),
                                array.get(14), array.get(15), array.get(16),
                                array.get(17)));

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
            menu.getItems().add(openCaseDetails);
            tableCases.setContextMenu(menu);


            // Selecting and Copy the Case Number to Clipboard
            tableCases.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {
                        copyCaseNumberToClipboard(tableCases);
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Get Case Number Failed", e);
                    }
                }
            });

            openCaseDetails.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    myCaseDetails();
                }
            });

            openCaseSFDC.setOnAction(new EventHandler<ActionEvent>() {
                @Override
                public void handle(ActionEvent event) {
                    try {

                        String search = "https://rbbn.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=001&sen=500&sen=005&sen=a0U&sen=00O&str="+getCaseNumber(tableCases, caseno);

                        URL caseSearch = new URL(search);
                        Desktop.getDesktop().browse(caseSearch.toURI());
                    }catch (Exception e){
                        logger.log(Level.WARNING, "Search Case in SFDC Failed!", e);
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
            logger.log(Level.WARNING, "Create Table Failed!", e);
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
        headerArray.add("Customer Interface");
        headerArray.add("Responsible");
        headerArray.add("RBBN Age");
        headerArray.add("Open Days");
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

        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short)10);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setWrapText(true);

        HSSFCellStyle style1 = workbook.createCellStyle();
        Font font1 = workbook.createFont();
        font1.setFontHeightInPoints((short)9);
        font1.setColor(IndexedColors.BLACK.getIndex());
        style1.setFont(font1);

        for (int k = 0; k < tableView.getColumns().size(); k++) {
            row.createCell(k).setCellValue(headerArray.get(k).toString());
            row.getCell(k).setCellStyle(style);
            spreadsheet.autoSizeColumn(k);
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

    private void extractToExcelProjects(TableView tableView, String textData, File file) throws IOException {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet spreadsheet = workbook.createSheet(textData);
        HSSFRow row = spreadsheet.createRow(0);
        TableColumn tableColumn = new TableColumn();

        ArrayList headerArray = new ArrayList();
        headerArray.add("Case No.");
        headerArray.add("Account Name");
        headerArray.add("Support Product");
        headerArray.add("Subject");
        headerArray.add("Last Modified Date");
        headerArray.add("Severity");
        headerArray.add("Status");
        headerArray.add("Case Number");
        headerArray.add("Support Hot List Reason");
        headerArray.add("Project Gating Date");
        headerArray.add("Support Theater");
        headerArray.add("Site Status");

        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short)10);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setWrapText(true);

        HSSFCellStyle style1 = workbook.createCellStyle();
        Font font1 = workbook.createFont();
        font1.setFontHeightInPoints((short)9);
        font1.setColor(IndexedColors.BLACK.getIndex());
        style1.setFont(font1);


        for (int k = 0; k < tableView.getColumns().size(); k++) {
            row.createCell(k).setCellValue(headerArray.get(k).toString());
            row.getCell(k).setCellStyle(style);
            spreadsheet.autoSizeColumn(k);
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
                row.getCell(j).setCellStyle(style1);
                spreadsheet.autoSizeColumn(j);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(file);
        workbook.write(fileOut);
        fileOut.close();
    }

    private void parseProjectDetailsData(){

        try {

            File csvfile = new File(dataFolder + "\\cmt_project_details.csv");
            HSSFWorkbook workBook = new HSSFWorkbook();
            String xlsFileAddress = dataFolder + "\\cmt_project_details.xls";
            HSSFSheet sheet = workBook.createSheet("Project Details");
            CreationHelper helper = workBook.getCreationHelper();

            int r = 0;

            CsvParserSettings settings = new CsvParserSettings();
            settings.setMaxCharsPerColumn(100000);
            settings.getFormat().setLineSeparator("\n");

            CsvParser parser = new CsvParser(settings);
            parser.beginParsing(csvfile);

            String[] row;
            while ((row = parser.parseNext()) != null) {

                Row frow = sheet.createRow((short) r++);
                for (int i = 0; i <row.length ; i++) {
                    frow.createCell(i).setCellValue(helper.createRichTextString(row[i]));
                }
            }

            parser.stopParsing();

            int lastRow = sheet.getLastRowNum();

            for (int i = 0; i < 5; i++) {
                sheet.removeRow(sheet.getRow(lastRow - i));
            }

            FileOutputStream fileOutputStream = new FileOutputStream(xlsFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();

        }catch (Exception e) {
            e.printStackTrace();
        }
    }



    private void accountsPage(){

        HSSFCell account;
        HSSFCell hotList;
        HSSFCell outFollow;
        HSSFCell escCases;
        HSSFCell caseSev;
        HSSFCell caseStat;
        HSSFCell ageCase;
        HSSFCell curResp;
        HSSFCell caseUpdate;
        HSSFCell caseOwner;
        HSSFCell rbnDays;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {
            HSSFSheet filtersheet = workbook.getSheetAt(0);

            accHotList = 0;
            accOutFollow = 0;
            accEscCases = 0;
            accBCCases = 0;
            accInactiveCases = 0;
            accBCDueCases = 0;
            accBCMissedCases = 0;
            accBCDSCases = 0;
            accBCInactiveCases = 0;
            accBCWIP = 0;
            accMJDueCases = 0;
            accMJMissedCases = 0;
            accRTSQueue = 0;
            accGPSQueue = 0;
            accMJUpdated = 0;
            accMJDSCases = 0;
            accMJWIP = 0;
            accQueuedCases = 0;
            accCoOwnerQueueCases = 0;
            accCoOwnerQueueCasesAssigned = 0;
            accE1Case = 0;
            accE2Cases = 0;
            accBCupdated = 0;
            accBCWac = 0;
            accMJWAC = 0;
            accMJInactiveCases = 0;
            accWOHCases = 0;
            accUpdateToday = 0;
            accUpdateMissed = 0;
            accUpdateNull = 0;
            accCoOwnCase = 0;
            accCoOwnQueue = 0;
            accBCQueue = 0;
            accMJQueue = 0;
            accMNQueue = 0;
            accMNWIP = 0;
            accMNWAC = 0;
            accMNUpdated = 0;
            accMNEng = 0;
            accMNInact = 0;
            accMNDue = 0;
            accMNMissedCases = 0;

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
                    case ("Account Name"):
                        caseAccountRef = i;
                        break;
                    case ("Days by Ribbon (days)"):
                        rbnDaysRefCell = i;
                        break;
                }
            }

            if (!txAccounts.getText().isEmpty()) {

                ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));
                int accfiltnum = setAcc.size();

                lblStatus.setText("ACCOUNT(S) VIEW:" + txAccounts.getText() + ")");

                if ((!setAcc.isEmpty())) {

                    for (int j = 0; j < accfiltnum; j++) {

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

                            hotList = filtersheet.getRow(i).getCell(caseHotListRefCell);
                            String strHotList = hotList.getStringCellValue();

                            outFollow = filtersheet.getRow(i).getCell(caseOutFolRefCell);
                            String followOut = outFollow.getStringCellValue();

                            account = filtersheet.getRow(i).getCell(caseAccountRef);
                            String acc = account.getStringCellValue();

                            ageCase = filtersheet.getRow(i).getCell(caseAgeRefCell);
                            String caseAge = ageCase.getStringCellValue();
                            int ageCaseNum = Integer.parseInt(caseAge);

                            rbnDays = filtersheet.getRow(i).getCell(rbnDaysRefCell);
                            String ribdays = rbnDays.getStringCellValue();
                            int ribDays = Integer.parseInt(ribdays);

                            caseUpdate = filtersheet.getRow(i).getCell(caseNextUpdateDateRef);
                            String caseupdate = caseUpdate.getStringCellValue();

                            LocalDate dateToday = LocalDate.now();
                            LocalDate caseUpdateDate = null;

                            if (acc.equals(setAcc.get(j))) {

                                if (!caseupdate.equals("NotSet")) {

                                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                                    caseUpdateDate = LocalDate.parse(caseupdate, formatter);
                                }

                                if (!strHotList.equals("NotSet") && !strHotList.equals("FALSE") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    accHotList++;
                                }

                                if (caseOwn.startsWith("PS")) {
                                    accGPSQueue++;
                                }

                                if (caseOwn.startsWith("TS") || caseOwn.startsWith("Tech-Ops")) {
                                    accRTSQueue++;
                                }

                                if (followOut.equals("1") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    accOutFollow++;
                                }
                                if (!escalatedCases.equals("NotSet") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    accEscCases++;
                                }

                                if (caseSever.equals("Critical") && !caseStatus.equals("Pending Closure")) {
                                    accE1Case++;
                                }

                                if (caseSever.equals("E2") && !caseStatus.equals("Pending Closure")) {
                                    accE2Cases++;
                                }

                                if (caseSever.equals("Business Critical")) {
                                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                        if (!caseStatus.equals("Develop Solution")) {
                                            if (ribDays <= 5) {
                                                accBCDueCases++;
                                            }
                                            if (ribDays > 5) {
                                                accBCMissedCases++;
                                            }
                                        } else {
                                            accBCDSCases++;
                                        }
                                        accBCCases++;
                                    } else {
                                        accBCInactiveCases++;
                                    }
                                    if (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault"))) {
                                        accBCWIP++;
                                    }
                                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                        if (responsible.equals("Customer action")) {
                                            accBCWac++;
                                        }
                                        if (responsible.equals("Customer updated")) {
                                            accBCupdated++;
                                        }
                                    }
                                    if (caseOwn.startsWith("PS") || caseOwn.startsWith("TS") || caseOwn.startsWith("Tech-Ops")){
                                        accBCQueue++;
                                    }
                                }

                                if (caseSever.equals("Major")) {
                                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                        if (!caseStatus.equals("Develop Solution")) {
                                            if (ribDays <= 15) {
                                                accMJDueCases++;
                                            }
                                            if (ribDays > 15) {
                                                accMJMissedCases++;
                                            }
                                        } else {
                                            accMJDSCases++;
                                        }
                                    } else {
                                        accMJInactiveCases++;
                                    }
                                    if (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault"))) {
                                        accMJWIP++;
                                    }
                                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                        if (responsible.equals("Customer action")) {
                                            accMJWAC++;
                                        }
                                        if (responsible.equals("Customer updated")) {
                                            accMJUpdated++;
                                        }
                                    }
                                    if (caseOwn.startsWith("PS") || caseOwn.startsWith("TS") || caseOwn.startsWith("Tech-Ops")){
                                        accMJQueue++;
                                    }
                                }
                                if (caseSever.equals("Minor")) {
                                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                        if (!caseStatus.equals("Develop Solution")) {
                                            if (ribDays <= 30) {
                                                accMNDue++;
                                            }
                                            if (ribDays > 30) {
                                                accMNMissedCases++;
                                            }
                                        } else {
                                            accMNEng++;
                                        }
                                    } else {
                                        accMNInact++;
                                    }
                                    if (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault"))) {
                                        accMNWIP++;
                                    }
                                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                        if (responsible.equals("Customer action")) {
                                            accMNWAC++;
                                        }
                                        if (responsible.equals("Customer updated")) {
                                            accMNUpdated++;
                                        }
                                    }
                                    if (caseOwn.startsWith("PS") || caseOwn.startsWith("TS") || caseOwn.startsWith("Tech-Ops")){
                                        accMNQueue++;
                                    }
                                }

                                if (caseStatus.equals("Pending Closure") || caseStatus.equals("Future Availability")) {
                                    accInactiveCases++;
                                }
                                if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                                    accWOHCases++;
                                }

                                if ((caseUpdateDate != null)) {
                                    if (caseUpdateDate.compareTo(dateToday) == 0) {
                                        accUpdateToday++;
                                    }
                                    if (caseUpdateDate.compareTo(dateToday) < 0) {
                                        accUpdateMissed++;
                                    }
                                }

                                if (caseupdate.equals("NotSet") && !caseStatus.equals("Pending Closure")) {
                                    accUpdateNull++;
                                }
                            }
                        }
                    }
                    /* Updating the button text from digested data */
                    btnAccHotIssues.setText(String.valueOf(accHotList));
                    btnAccOutFollow.setText(String.valueOf(this.accOutFollow));
                    btnAccEscalated.setText(String.valueOf(accEscCases));
                    btnAccBCCases.setText(String.valueOf(accBCCases));
                    btnAccInactive.setText(String.valueOf(accInactiveCases));
                    btnAccBCDue.setText(String.valueOf(accBCDueCases));
                    btnAccBCMissed.setText(String.valueOf(accBCMissedCases));
                    btnAccBCWac.setText(String.valueOf(accBCWac));
                    btnAccBCupdated.setText(String.valueOf(accBCupdated));
                    btnAccBCEngineering.setText(String.valueOf(accBCDSCases));
                    btnAccBCINACT.setText(String.valueOf(accBCInactiveCases));
                    btnAccBCWIP.setText(String.valueOf(accBCWIP));
                    btnAccMJDue.setText(String.valueOf(accMJDueCases));
                    btnAccMJMissed.setText(String.valueOf(accMJMissedCases));
                    btnAccMJWac.setText(String.valueOf(accMJWAC));
                    btnAccMJupdated.setText(String.valueOf(accMJUpdated));
                    btnAccMJEngineering.setText(String.valueOf(accMJDSCases));
                    btnAccMJINACT.setText(String.valueOf(accMJInactiveCases));
                    btnAccMJWIP.setText(String.valueOf(accMJWIP));
                    btnaccGPSQueue.setText(String.valueOf(accGPSQueue));
                    btnAccRTSQueue.setText(String.valueOf(accRTSQueue));
                    btnAccE1Cases.setText(String.valueOf(accE1Case));
                    btnAccE2Cases.setText(String.valueOf(accE2Cases));
                    btnAccWOH.setText(String.valueOf(accWOHCases));
                    btnAccUpdateToday.setText(String.valueOf(accUpdateToday));
                    btnAccUpdateMissed.setText(String.valueOf(accUpdateMissed));
                    btnAccUpdateNull.setText(String.valueOf(accUpdateNull));
                    btnAccMNMissed.setText(String.valueOf(accMNMissedCases));
                    btnAccBCQueue.setText(String.valueOf(accBCQueue));
                    btnAccMJQueue.setText(String.valueOf(accMJQueue));
                    btnAccMNQueue.setText(String.valueOf(accMNQueue));
                    btnAccMNWIP.setText(String.valueOf(accMNWIP));
                    btnAccMNWac.setText(String.valueOf(accMNWAC));
                    btnAccMNupdated.setText(String.valueOf(accMNUpdated));
                    btnAccMNEngineering.setText(String.valueOf(accMNEng));
                    btnAccMNINACT.setText(String.valueOf(accMNInact));
                    btnAccMNDue.setText(String.valueOf(accMNDue));

                }
            }

        }catch (Exception e)
        {
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
        HSSFCell rbnDays;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {
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
            dueMNday = 0;
            misMJdue = 0;
            misMNdue = 0;
            custActMJ = 0;
            custRpdMJ = 0;
            custRpdMN = 0;
            custActMN = 0;
            MJds = 0;
            MJpc = 0;
            MNpc = 0;
            MJwip = 0;
            MNwip = 0;
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
                    case ("Days by Ribbon (days)"):
                        rbnDaysRefCell = i;
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

                rbnDays = filtersheet.getRow(i).getCell(rbnDaysRefCell);
                String rbbnAge = rbnDays.getStringCellValue();
                int ribDays = Integer.parseInt(rbbnAge);

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

                if (caseOwn.startsWith("PS") || caseOwn.startsWith("TS") || caseOwn.startsWith("Tech-Ops")){
                    switch (caseSever) {
                        case ("Business Critical"):
                            bcQueue++;
                            break;
                        case ("Major"):
                            mjQueue++;
                            break;
                        case ("Minor"):
                            mnQueue++;
                            break;
                    }
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
                            if (ribDays <= 5) {
                                bcDue++;
                            }
                            if (ribDays > 5) {
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
                            if (ribDays <= 15) {
                                dueMJday++;
                            }
                            if (ribDays > 15) {
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
                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        if (!caseStatus.equals("Develop Solution")) {
                            if (ribDays <= 30) {
                                dueMNday++;
                            }
                            if (ribDays > 30) {
                                misMNdue++;
                            }
                        } else {
                            MNds++;
                        }
                    } else {
                        MNpc++;
                    }
                    if (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault"))) {
                        MNwip++;
                    }
                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        if (responsible.equals("Customer action")) {
                            custActMN++;
                        }
                        if (responsible.equals("Customer updated")) {
                            custRpdMN++;
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
            btnMNWIP.setText(String.valueOf(MNwip));
            btnMNWac.setText(String.valueOf(custActMN));
            btnMNupdated.setText(String.valueOf(custRpdMN));
            btnMNEngineering.setText(String.valueOf(MNds));
            btnMNINACT.setText(String.valueOf(MNpc));
            btnMNDue.setText(String.valueOf(dueMNday));
            btnPSQueue.setText(String.valueOf(queuePS));
            btnTSQueue.setText(String.valueOf(queueTS));
            btnE1Cases.setText(String.valueOf(e1Cases));
            btnE2Cases.setText(String.valueOf(e2Cases));
            btnWOH.setText(String.valueOf(wohCases));
            btnUpdateToday.setText(String.valueOf(updateToday));
            btnUpdateMissed.setText(String.valueOf(updateMissed));
            btnUpdateNull.setText(String.valueOf(updateNull));
            btnMNMissed.setText(String.valueOf(misMNdue));
            btnBCQueue.setText(String.valueOf(bcQueue));
            btnMJQueue.setText(String.valueOf(mjQueue));
            btnMNQueue.setText(String.valueOf(mnQueue));

            /* Updating completed for overview page */

        } catch (Exception e) {
            logger.log(Level.WARNING, "Unable To Build Overview Page... ", e);
        }
    }

    private void regionCases(){

        HSSFCell region;
        HSSFCell hotList;
        HSSFCell outFollow;
        HSSFCell escCases;
        HSSFCell caseSev;
        HSSFCell caseStat;
        HSSFCell ageCase;
        HSSFCell curResp;
        HSSFCell caseUpdate;
        HSSFCell caseOwner;
        HSSFCell rbnAge;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();

            regHotList = 0;
            regOutFollow = 0;
            regEscCases = 0;
            regBCCases = 0;
            regInactiveCases = 0;
            regBCDueCases = 0;
            regBCMissedCases = 0;
            regBCDSCases = 0;
            regBCInactiveCases = 0;
            regBCWIP = 0;
            regMJDueCases = 0;
            regMJMissedCases = 0;
            regMNMissedCases = 0;
            regRTSQueue = 0;
            regGPSQueue = 0;
            regMJUpdated = 0;
            regMJDSCases = 0;
            regMJWIP = 0;
            regQueuedCases = 0;
            regCoOwnerQueueCases = 0;
            regCoOwnerQueueCasesAssigned = 0;
            regE1Case = 0;
            regE2Cases = 0;
            regBCupdated = 0;
            regBCWac = 0;
            regMJWAC = 0;
            regMJInactiveCases = 0;
            regWOHCases = 0;
            regUpdateToday = 0;
            regUpdateMissed = 0;
            regUpdateNull = 0;
            regCoOwnCase = 0;
            regCoOwnQueue = 0;
            regBCQueue = 0;
            regMJQueue = 0;
            regMNQueue = 0;
            regMNWIP = 0;
            regMNWAC = 0;
            regMNUpdated = 0;
            regMNEng = 0;
            regMNInact = 0;
            regMNDue = 0;
            regMNMissed = 0;

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
                    case ("Support Theater"):
                        caseRegionRef = i;
                        break;
                    case ("Days by Ribbon (days)"):
                        rbnDaysRegRefCell = i;
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

                hotList = filtersheet.getRow(i).getCell(caseHotListRefCell);
                String strHotList = hotList.getStringCellValue();

                outFollow = filtersheet.getRow(i).getCell(caseOutFolRefCell);
                String followOut = outFollow.getStringCellValue();

                region = filtersheet.getRow(i).getCell(caseRegionRef);
                String reg = region.getStringCellValue();

                ageCase = filtersheet.getRow(i).getCell(caseAgeRefCell);
                String caseAge = ageCase.getStringCellValue();
                int ageCaseNum = Integer.parseInt(caseAge);

                rbnAge = filtersheet.getRow(i).getCell(rbnDaysRegRefCell);
                String rbnDays = rbnAge.getStringCellValue();
                int ribDays = Integer.parseInt(rbnDays);

                caseUpdate = filtersheet.getRow(i).getCell(caseNextUpdateDateRef);
                String caseupdate = caseUpdate.getStringCellValue();

                LocalDate dateToday = LocalDate.now();
                LocalDate caseUpdateDate = null;

                if (reg.equals(selectedRegion)) {

                    if (!caseupdate.equals("NotSet")) {

                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
                        caseUpdateDate = LocalDate.parse(caseupdate, formatter);
                    }

                    if (!strHotList.equals("NotSet") && !strHotList.equals("FALSE") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        regHotList++;
                    }

                    if (caseOwn.startsWith("PS")) {
                        regGPSQueue++;
                    }

                    if (caseOwn.startsWith("TS") || caseOwn.startsWith("Tech-Ops")) {
                        regRTSQueue++;
                    }

                    if(caseOwn.startsWith("PS") || caseOwn.startsWith("TS") || caseOwn.startsWith("Tech-Ops")){
                        if (caseSever.equals("Business Critical")){
                            regBCQueue++;
                        }
                        if (caseSever.equals("Major")){
                            regMJQueue++;
                        }
                        if (caseSever.equals("Minor")){
                            regMNQueue++;
                        }
                    }

                    if (followOut.equals("1") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        regOutFollow++;
                    }
                    if (!escalatedCases.equals("NotSet") && !caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        regEscCases++;
                    }

                    if (caseSever.equals("Critical") && !caseStatus.equals("Pending Closure")) {
                        regE1Case++;
                    }

                    if (caseSever.equals("E2") && !caseStatus.equals("Pending Closure")) {
                        regE2Cases++;
                    }

                    if (caseSever.equals("Business Critical")) {
                        if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                            if (!caseStatus.equals("Develop Solution")) {
                                if (ribDays <= 5) {
                                    regBCDueCases++;
                                }
                                if (ribDays > 5) {
                                    regBCMissedCases++;
                                }
                            } else {
                                regBCDSCases++;
                            }
                            regBCCases++;
                        } else {
                            regBCInactiveCases++;
                        }
                        if (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault"))) {
                            regBCWIP++;
                        }
                        if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                            if (responsible.equals("Customer action")) {
                                regBCWac++;
                            }
                            if (responsible.equals("Customer updated")) {
                                regBCupdated++;
                            }
                        }
                    }

                    if (caseSever.equals("Major")) {
                        if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                            if (!caseStatus.equals("Develop Solution")) {
                                if (ribDays <= 15) {
                                    regMJDueCases++;
                                }
                                if (ribDays > 15) {
                                    regMJMissedCases++;
                                }
                            } else {
                                regMJDSCases++;
                            }
                        } else {
                            regMJInactiveCases++;
                        }
                        if (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault"))) {
                            regMJWIP++;
                        }
                        if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                            if (responsible.equals("Customer action")) {
                                regMJWAC++;
                            }
                            if (responsible.equals("Customer updated")) {
                                regMJUpdated++;
                            }
                        }
                    }
                    if (caseSever.equals("Minor")) {

                        if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                            if (!caseStatus.equals("Develop Solution")) {
                                if (ribDays <= 30) {
                                    regMNDue++;
                                }
                                if (ribDays > 30) {
                                    regMNMissed++;
                                }
                            } else {
                                regMNEng++;
                            }
                        } else {
                            regMNInact++;
                        }
                        if (caseStatus.equals("Open / Assign") || (caseStatus.equals("Isolate Fault"))) {
                            regMNWIP++;
                        }
                        if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                            if (responsible.equals("Customer action")) {
                                regMNWAC++;
                            }
                            if (responsible.equals("Customer updated")) {
                                regMNUpdated++;
                            }
                        }
                    }

                    if (caseStatus.equals("Pending Closure") || caseStatus.equals("Future Availability")) {
                        regInactiveCases++;
                    }
                    if (!caseStatus.equals("Pending Closure") && !caseStatus.equals("Future Availability")) {
                        regWOHCases++;
                    }

                    if ((caseUpdateDate != null)) {
                        if (caseUpdateDate.compareTo(dateToday) == 0) {
                            regUpdateToday++;
                        }
                        if (caseUpdateDate.compareTo(dateToday) < 0) {
                            regUpdateMissed++;
                        }
                    }

                    if (caseupdate.equals("NotSet") && !caseStatus.equals("Pending Closure")) {
                        regUpdateNull++;
                    }
                }

            }
        }catch(Exception e){
            e.printStackTrace();
        }

        btnRegE1Cases.setText(String.valueOf(regE1Case));
        btnRegE2Cases.setText(String.valueOf(regE2Cases));
        btnRegOutFollow.setText(String.valueOf(regOutFollow));
        btnRegEscalated.setText(String.valueOf(regEscCases));
        btnRegBCCases.setText(String.valueOf(regBCCases));
        btnRegHotIssues.setText(String.valueOf(regHotList));
        btnRegInactive.setText(String.valueOf(regInactiveCases));
        btnRegBCWIP.setText(String.valueOf(regBCWIP));
        btnRegBCWac.setText(String.valueOf(regBCWac));
        btnRegBCupdated.setText(String.valueOf(regBCupdated));
        btnRegBCEngineering.setText(String.valueOf(regBCDSCases));
        btnRegBCINACT.setText(String.valueOf(regBCInactiveCases));
        btnRegMJWIP.setText(String.valueOf(regMJWIP));
        btnRegMJWac.setText(String.valueOf(regMJWAC));
        btnRegMJupdated.setText(String.valueOf(regMJUpdated));
        btnRegMJEngineering.setText(String.valueOf(regMJDSCases));
        btnRegMJINACT.setText(String.valueOf(regMJInactiveCases));
        btnRegBCDue.setText(String.valueOf(regBCDueCases));
        btnRegBCMissed.setText(String.valueOf(regBCMissedCases));
        btnRegMJDue.setText(String.valueOf(regMJDueCases));
        btnRegMJMissed.setText(String.valueOf(regMJMissedCases));
        btnRegWOH.setText(String.valueOf(regWOHCases));
        btnRegUpdateToday.setText(String.valueOf(regUpdateToday));
        btnRegUpdateMissed.setText(String.valueOf(regUpdateMissed));
        btnRegUpdateNull.setText(String.valueOf(regUpdateNull));
        btnRegMNMissed.setText(String.valueOf(regMNMissed));
        btnRegRTSQueue.setText(String.valueOf(regRTSQueue));
        btnRegGPSQueue.setText(String.valueOf(regGPSQueue));
        btnRegBCQueue.setText(String.valueOf(regBCQueue));
        btnRegMJQueue.setText(String.valueOf(regMJQueue));
        btnRegMNQueue.setText(String.valueOf(regMNQueue));
        btnRegMNWIP.setText(String.valueOf(regMNWIP));
        btnRegMNWac.setText(String.valueOf(regMNWAC));
        btnRegMNupdated.setText(String.valueOf(regMNUpdated));
        btnRegMNINACT.setText(String.valueOf(regMNInact));
        btnRegMNDue.setText(String.valueOf(regMNDue));
        btnRegMNEngineering.setText(String.valueOf(regMNEng));
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
        HSSFCell myCoOwnedCase;
        HSSFCell myCoOwnerQueue;
        HSSFCell myRbbnDays;
        HSSFCell myUser;
        HSSFCell mySev;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
            myCoOwnerQueueCases = 0;
            myCoOwnerQueueCasesAssigned = 0;
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
            myCoOwnCase = 0;
            myCoOwnQueue = 0;
            myBCInQueue = 0;
            myMJInQueue = 0;
            myMNInQueue = 0;
            myMNUpdated = 0;
            myMNWAC = 0;
            myMNWIP = 0;
            myMNEng = 0;
            myMNDue = 0;
            myMNMissed = 0;
            myMNINAct = 0;

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
                    case ("Co-Owner"):
                        myCoOwnCaseRefCell = i;
                        break;
                    case ("Co-Owner Queue"):
                        myCoOwnQueueRefCell = i;
                        break;
                    case ("Days by Ribbon (days)"):
                         myRbbnDaysRefCell= i;
                        break;
                }
            }

            /* Creating Input Data Arrays from Setttings Page */

            wgToNames();

            if (!txUsers.getText().isEmpty()||!txWorkGroup.getText().isEmpty()) {

                ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));

                if (!txWorkGroup.getText().isEmpty()){
                    setWgNames.removeAll(setUser);
                    setUser.addAll(setWgNames);
                }

                int userfiltnum = setUser.size();

                if ((!setUser.isEmpty())) {

                    for (int j = 0; j < userfiltnum; j++) {

                        for (int i = 1; i < lastRow + 1; i++) {

                            caseUser = filtersheet.getRow(i).getCell(mycaseOwnerRefCell);
                            String caseuser = caseUser.getStringCellValue();

                            myCoOwnedCase = filtersheet.getRow(i).getCell(myCoOwnCaseRefCell);
                            String coOwner = myCoOwnedCase.getStringCellValue();

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

                            myRbbnDays = filtersheet.getRow(i).getCell(myRbbnDaysRefCell);
                            String myRbnDays = myRbbnDays.getStringCellValue();
                            int myRibDays = Integer.parseInt(myRbnDays);

                            if (caseuser.equals(setUser.get(j)) || coOwner.equals(setUser.get(j))) {

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
                                        if (myRibDays <= 5) {
                                            myBCDueCases++;
                                        }
                                        if (myRibDays > 5) {
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
                                        if (myRibDays <= 15) {
                                            myMJDueCases++;
                                        }
                                        if (myRibDays > 15) {
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
                                if (mycaseSever.equals("Minor")){
                                    if (mycaseStatus.equals("Develop Solution")){
                                        myMNEng++;
                                    }
                                    if (mycaseStatus.equals("Open / Assign") || (mycaseStatus.equals("Isolate Fault"))){
                                        myMNWIP++;
                                        if (myRibDays <= 30){
                                            myMNDue++;
                                        }
                                        if (myRibDays > 30){
                                            myMNMissed++;
                                        }
                                    }
                                    if (mycaseStat.equals("Pending Closure") || mycaseStatus.equals("Future Availability")){
                                        myMNINAct++;
                                    }
                                    if (myresponsible.equals("Customer action")){
                                        myMNWAC++;
                                    }
                                    if (myresponsible.equals("Customer updated")){
                                        myMNUpdated++;
                                    }
                                }

                                if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                    myInactiveCases++;
                                } else {
                                    myWOHCases++;
                                }

                                if ((caseUpdateDate != null)) {

                                    if(!myCaseUpdate.equals("NotSet")) {

                                        if (caseUpdateDate.compareTo(dateToday) == 0) {
                                            myUpdateToday++;
                                        }
                                        if (caseUpdateDate.compareTo(dateToday) < 0) {
                                            myUpdateMissed++;
                                        }
                                    }
                                }
                                if (myCaseUpdate.equals("NotSet")) {
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

                            myCoOwnerQueue = workbook.getSheetAt(0).getRow(l).getCell(myCoOwnQueueRefCell);
                            String coOwnQueue = myCoOwnerQueue.getStringCellValue();

                            myCoOwnedCase = workbook.getSheetAt(0).getRow(l).getCell(myCoOwnCaseRefCell);
                            String coOwn = myCoOwnedCase.getStringCellValue();

                            mySev = workbook.getSheetAt(0).getRow(l).getCell(mycaseSevRefCell);
                            String mySever = mySev.getStringCellValue();

                            if (casequeue.equals(setQueue.get(k))){
                                myQueuedCases++;
                            }

                            if (coOwnQueue.equals(setQueue.get(k)) && coOwn.equals("NotSet")){
                                myCoOwnerQueueCases++;

                            }
                            if (coOwnQueue.equals(setQueue.get(k)) && !coOwn.equals("NotSet")){
                                myCoOwnerQueueCasesAssigned++;
                            }
                            if ((casequeue.equals(setQueue.get(k))) && (casequeue.startsWith("PS") || casequeue.startsWith("TS") || casequeue.startsWith("Tech-Ops"))){
                                switch (mySever) {
                                    case ("Business Critical"):
                                        myBCInQueue++;
                                        break;
                                    case ("Major"):
                                        myMJInQueue++;
                                        break;
                                    case ("Minor"):
                                        myMNInQueue++;
                                        break;
                                }
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
            //btnMyCoOwnQueue.setText(String.valueOf(myCoOwnerQueueCases));
            //btnMyCoQueueAssigned.setText(String.valueOf(myCoOwnerQueueCasesAssigned));
            btnMyBCQueue.setText(String.valueOf(myBCInQueue));
            btnMyMJQueue.setText(String.valueOf(myMJInQueue));
            btnMyMNQueue.setText(String.valueOf(myMNInQueue));
            btnMyMNWIP.setText(String.valueOf(myMNWIP));
            btnMyMNWac.setText(String.valueOf(myMNWAC));
            btnMyMNupdated.setText(String.valueOf(myMNUpdated));
            btnMyMNEngineering.setText(String.valueOf(myMNEng));
            btnMyMNINACT.setText(String.valueOf(myMNINAct));
            btnMyMNDue.setText(String.valueOf(myMNDue));
            btnMyMNMissed.setText(String.valueOf(myMNMissed));

        } catch (Exception e) {
            logger.log(Level.WARNING, "Unable To Build My Cases Page...", e);        }
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
        HSSFCell caseRbnDays;
        HSSFCell caseRegion;


        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
            prodBCQueue = 0;
            prodMJQueue = 0;
            prodMNQueue = 0;
            prodMNWIP = 0;
            prodMNWAC = 0;
            prodMNUpdated = 0;
            prodMNInact = 0;
            prodMNEng = 0;
            prodMNDue = 0;
            prodMNMissed = 0;

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
                    case ("Days by Ribbon (days)"):
                        rbbnDaysRefCell = i;
                        break;
                    case ("Support Theater"):
                        caseRegionRef = i;
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

                            caseRbnDays = filtersheet.getRow(i).getCell(rbnDaysRefCell);
                            String prodRbnDays = caseRbnDays.getStringCellValue();
                            int prdRbnDays = Integer.parseInt(prodRbnDays);

                            caseRegion = filtersheet.getRow(i).getCell(caseRegionRef);
                            String productRegion = caseRegion.getStringCellValue();


                            if (reg_product){
                                if ((productCellStr.equals(setProducts.get(j))) && regChoice.getValue().equals(productRegion)) {

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
                                        if (!mycaseStatus.equals("Develop Solution") || !mycaseStatus.equals("Future Availability") || !mycaseStatus.equals("Pending Closure")) {
                                            if (prdRbnDays <= 5) {
                                                prodBCDueCases++;
                                            }
                                            if (prdRbnDays > 5) {
                                                prodBCMissedCases++;
                                            }
                                        }
                                        if (mycaseStatus.equals("Develop Solution")) {
                                            prodBCDSCases++;
                                        }
                                        if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                            prodBCInactiveCases++;
                                        }
                                        if (caseuser.startsWith("PS") || caseuser.startsWith("TS") || caseuser.startsWith("Tech-Ops")) {
                                            prodBCQueue++;
                                        }
                                    }
                                    if (mycaseSever.equals("Major")) {

                                        if (mycaseStatus.equals("Develop Solution")) {
                                            prodMJDSCases++;
                                        }
                                        if (!mycaseStatus.equals("Develop Solution") || !mycaseStatus.equals("Future Availability") || !mycaseStatus.equals("Pending Closure")) {
                                            if (prdRbnDays <= 15) {
                                                prodMJDueCases++;
                                            }
                                            if (prdRbnDays > 15) {
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
                                        if (caseuser.startsWith("PS") || caseuser.startsWith("TS") || caseuser.startsWith("Tech-Ops")) {
                                            prodMJQueue++;
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
                                    if (mycaseSever.equals("Minor")) {
                                        if (mycaseStatus.equals("Develop Solution")) {
                                            prodMNEng++;
                                        }
                                        if (!mycaseStatus.equals("Develop Solution") || !mycaseStatus.equals("Future Availability") || !mycaseStatus.equals("Pending Closure")) {
                                            if (prdRbnDays <= 30) {
                                                prodMNDue++;
                                            }
                                            if (prdRbnDays > 30) {
                                                prodMNMissed++;
                                            }
                                        }
                                        if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                            prodMNInact++;
                                        }
                                        if (myresponsible.equals("Customer action")) {
                                            prodMNWAC++;
                                        }
                                        if (myresponsible.equals("Customer updated")) {
                                            prodMNUpdated++;
                                        }
                                        if (mycaseStatus.equals("Open / Assign") || (mycaseStatus.equals("Isolate Fault"))) {
                                            prodMNWIP++;
                                        }
                                        if (caseuser.startsWith("PS") || caseuser.startsWith("TS") || caseuser.startsWith("Tech-Ops")) {
                                            prodMNQueue++;
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
                            else{
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
                                        if (!mycaseStatus.equals("Develop Solution") || !mycaseStatus.equals("Future Availability") || !mycaseStatus.equals("Pending Closure")) {
                                            if (prdRbnDays <= 5) {
                                                prodBCDueCases++;
                                            }
                                            if (prdRbnDays > 5) {
                                                prodBCMissedCases++;
                                            }
                                        }
                                        if (mycaseStatus.equals("Develop Solution")) {
                                            prodBCDSCases++;
                                        }
                                        if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                            prodBCInactiveCases++;
                                        }
                                        if (caseuser.startsWith("PS") || caseuser.startsWith("TS") || caseuser.startsWith("Tech-Ops")) {
                                            prodBCQueue++;
                                        }
                                    }
                                    if (mycaseSever.equals("Major")) {

                                        if (mycaseStatus.equals("Develop Solution")) {
                                            prodMJDSCases++;
                                        }
                                        if (!mycaseStatus.equals("Develop Solution") || !mycaseStatus.equals("Future Availability") || !mycaseStatus.equals("Pending Closure")) {
                                            if (prdRbnDays <= 15) {
                                                prodMJDueCases++;
                                            }
                                            if (prdRbnDays > 15) {
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
                                        if (caseuser.startsWith("PS") || caseuser.startsWith("TS") || caseuser.startsWith("Tech-Ops")) {
                                            prodMJQueue++;
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
                                    if (mycaseSever.equals("Minor")) {
                                        if (mycaseStatus.equals("Develop Solution")) {
                                            prodMNEng++;
                                        }
                                        if (!mycaseStatus.equals("Develop Solution") || !mycaseStatus.equals("Future Availability") || !mycaseStatus.equals("Pending Closure")) {
                                            if (prdRbnDays <= 30) {
                                                prodMNDue++;
                                            }
                                            if (prdRbnDays > 30) {
                                                prodMNMissed++;
                                            }
                                        }
                                        if (mycaseStatus.equals("Pending Closure") || mycaseStatus.equals("Future Availability")) {
                                            prodMNInact++;
                                        }
                                        if (myresponsible.equals("Customer action")) {
                                            prodMNWAC++;
                                        }
                                        if (myresponsible.equals("Customer updated")) {
                                            prodMNUpdated++;
                                        }
                                        if (mycaseStatus.equals("Open / Assign") || (mycaseStatus.equals("Isolate Fault"))) {
                                            prodMNWIP++;
                                        }
                                        if (caseuser.startsWith("PS") || caseuser.startsWith("TS") || caseuser.startsWith("Tech-Ops")) {
                                            prodMNQueue++;
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
            }

            if (!txQueues.getText().isEmpty()) {

                ArrayList<String> setQueue = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));
                int queuefiltnum = setQueue.size();

                if (!setQueue.isEmpty()) {

                    for (int k = 0; k < queuefiltnum; k++) {

                        for (int l = 0; l < lastRow + 1; l++) {

                            caseQueue = workbook.getSheetAt(0).getRow(l).getCell(caseOwnerRefCell);
                            String casequeue = caseQueue.getStringCellValue();

                            caseRegion = workbook.getSheetAt(0).getRow(l).getCell(caseRegionRef);
                            String productRegion = caseRegion.getStringCellValue();

                            if (reg_product){
                                if (casequeue.equals(setQueue.get(k)) && regChoice.getValue().equals(productRegion)) {
                                    prodQueuedCases++;
                                }
                            }
                            else{
                                if (casequeue.equals(setQueue.get(k))) {
                                    prodQueuedCases++;
                                }
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
            btnBCQueueProd.setText(String.valueOf(prodBCQueue));
            btnMJQueueProd.setText(String.valueOf(prodMJQueue));
            btnMNQueueProd.setText(String.valueOf(prodMNQueue));
            btnMNWIPProd.setText(String.valueOf(prodMNWIP));
            btnMNWacProd.setText(String.valueOf(prodMNWAC));
            btnMNupdatedProd.setText(String.valueOf(prodMNUpdated));
            btnMNEngineeringProd.setText(String.valueOf(prodMNEng));
            btnMNINACTProd.setText(String.valueOf(prodMNInact));
            btnMNDueProd.setText(String.valueOf(prodMNDue));
            btnMNMissedProd.setText(String.valueOf(prodMNMissed));


        } catch (Exception e) {
            logger.log(Level.WARNING, "Unable To Build Products Page... ", e);
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
            casesTooltip.setText("Personalized View\n" +
                    "My Cases View");
            btnCases.setTooltip(casesTooltip);
        }

        if (event.getSource() == btnCustomers) {

            Tooltip customersTooltip = new Tooltip();
            customersTooltip.setText("Customer Based\n" +
                    "Case View");
            btnCustomers.setTooltip(customersTooltip);
        }

        if (event.getSource() == btnMyNotes) {

            Tooltip surveyTooltip = new Tooltip();
            surveyTooltip.setText("Personal Memo Book...");
            btnMyNotes.setTooltip(surveyTooltip);
        }

        if (event.getSource() == btnSettings) {

            Tooltip settingsTooltip = new Tooltip();
            settingsTooltip.setText("Customize\n" +
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
            userTextBoxTip.setText("Please select users from pick list");
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
            CoOwnerCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseCoOwner"));
            CoOwnerQueueCol.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseCoOwnerQueue"));
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
        }
        if (table == tableCustomers) {

            NumberColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseNumber"));
            SeverityColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseSeverity"));
            StatusColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseStatus"));
            OwnerColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseOwner"));
            CoOwnerColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseCoOwner"));
            CoOwnerQueuColCust.setCellValueFactory(new PropertyValueFactory<CaseTableView, String>("caseCoOwnerQueue"));
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

        }

    }

    private void initProjectTable(){

        prjNoCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjCaseNo"));
        prjAccCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjCaseAccount"));
        prjProdCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjCaseProduct"));
        prjSubCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjCaseSubject"));
        prjModCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjModDate"));
        prjStatCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjCaseSeverity"));
        prjSevCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjCaseStatus"));
        prjHotRCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjHotR"));
        prjGateDateCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjGateDate"));
        prjRegionCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjRegion"));
        prjSiteStatusCol.setCellValueFactory(new PropertyValueFactory<ProjectTableView, String>("prjSiteStatus"));
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
        saveCaseDetails(caseview);
    }

    private void copyCaseNumberToClipboardProjects(TableView<ProjectTableView> tableCases) {

        TablePosition tablePosition = (TablePosition) tableCases.getSelectionModel().getSelectedCells().get(0);
        int row = tablePosition.getRow();
        ProjectTableView caseview = (ProjectTableView) tableCases.getItems().get(row);
        TableColumn tableColumn = tablePosition.getTableColumn();
        String data1 = caseview.getPrjCaseNo();
        ClipboardContent content = new ClipboardContent();
        content.putString(data1);
        Clipboard.getSystemClipboard().setContent(content);

        HSSFCell cellVal;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_projects.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Case No.")) {
                    caseNumCellRef = i;
                }
            }

            for (int k = 1; k < lastRow + 1; k++) {

                cellVal = filtersheet.getRow(k).getCell(caseNumCellRef);
                String cellValToCompare = cellVal.getStringCellValue();

                if (cellValToCompare.equals(data1)){

                    selectedCase = new ArrayList<>();
                    Iterator<org.apache.poi.ss.usermodel.Cell> iterCells = filtersheet.getRow(k).cellIterator();
                    while (iterCells.hasNext()) {
                        HSSFCell cell = (HSSFCell) iterCells.next();
                        selectedCase.add(cell.getStringCellValue());
                    }

                }
            }

        }catch (Exception e){
            logger.log(Level.WARNING, "Unable To Get Case Number...", e);
        }

        int selectedsize= selectedCase.size();

        try {

            File caseSelFile = new File(detailsFolder + "\\" + "caseSelProject");
            BufferedWriter br = new BufferedWriter(new FileWriter(caseSelFile));
            StringBuilder sb = new StringBuilder();

            // Append strings from array
            for (String element : selectedCase) {
                sb.append(element);
                sb.append("\",\"");
            }
            br.write(sb.toString());
            br.close();

            /*File caseSelFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\CaseDetails\\" + "caseSelProject");

            FileWriter writer = new FileWriter(caseSelFile);

            writer.write(String.valueOf(selectedCase));

            for (int i = 0; i <selectedsize ; i++) {

                writer.write(selectedCase.get(i) + "\n");
            }

            writer.close();*/

        }catch (Exception e){
            logger.log(Level.WARNING, "Unable To Save Case Details...", e);
        }
    }

    private String getCaseNumber(TableView<CaseTableView> tableCases, String caseNumber){

        TablePosition tablePosition = (TablePosition) tableCases.getSelectionModel().getSelectedCells().get(0);
        int row = tablePosition.getRow();
        CaseTableView caseview = (CaseTableView) tableCases.getItems().get(row);
        TableColumn tableColumn = tablePosition.getTableColumn();
        caseNumber = caseview.getCaseNumber();
        return caseNumber;
    }

    private String getCaseNumberProjects(TableView<ProjectTableView> tableCases, String caseNumber){

        TablePosition tablePosition = (TablePosition) tableCases.getSelectionModel().getSelectedCells().get(0);
        int row = tablePosition.getRow();
        ProjectTableView caseview = (ProjectTableView) tableCases.getItems().get(row);
        TableColumn tableColumn = tablePosition.getTableColumn();
        caseNumber = caseview.getPrjCaseNo();
        return caseNumber;
    }


    private void exportExcelAction(TableView<CaseTableView> table) {

        try {
            FileChooser fileChooser = new FileChooser();
            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("XLS files (*.xls)", "*.xls");
            try {
                fileChooser.setInitialDirectory(new File(rootFolder));
            } catch (Exception e){
                logger.log(Level.WARNING, "Unable To Export Table to Excel To Desktop, Using defaults...", e);
                fileChooser.setInitialDirectory(new File(rootFolder));
            }

            fileChooser.getExtensionFilters().add(extFilter);

            Stage primaryStage = new Stage();

            File file = fileChooser.showSaveDialog(primaryStage);
            primaryStage.show();

            if (file != null) {

                extractToExcel(table, "testData", file);
            }
            primaryStage.close();
        } catch (Exception e) {
            logger.log(Level.WARNING, "Unable To Export Table to Excel...", e);
        }

    }
    private void exportExcelActionProjects(TableView<ProjectTableView> table) {

        try {
            FileChooser fileChooser = new FileChooser();
            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("XLS files (*.xls)", "*.xls");
            try {
                fileChooser.setInitialDirectory(new File(rootFolder));
            } catch (Exception e){
                logger.log(Level.WARNING, "Unable To Export Table to Excel To Desktop, Using defaults...", e);
                fileChooser.setInitialDirectory(new File(rootFolder));
            }

            fileChooser.getExtensionFilters().add(extFilter);

            Stage primaryStage = new Stage();

            File file = fileChooser.showSaveDialog(primaryStage);
            primaryStage.show();

            if (file != null) {


                if (table != tableProjects) {
                    extractToExcel(table, "CMT", file);
                }else{
                    extractToExcelProjects(table, "CMT Projects", file);
                }
            }
            primaryStage.close();
        } catch (Exception e) {
            logger.log(Level.WARNING, "Unable To Export Project Table to Excel...", e);
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

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
            logger.log(Level.WARNING, "Unable To Build Customer Page...", e);
        }
    }

    private void alertNoComment(){

        Alert alert = new Alert(Alert.AlertType.WARNING);
        ((Stage) alert.getDialogPane().getScene().getWindow()).getIcons().add(new Image("home/image/rbbicon.png"));
        alert.setTitle("RBBN CMT WARNING:");
        alert.setHeaderText(null);
        alert.setContentText("THERE IS NO COMMENT IN THIS CASE..." + "\n" + "\n" + "CREATED IN LAST 7 DAYS!");
        alert.showAndWait();

    }

    private void userSelectArray() {

        HSSFCell userCell;
        HSSFCell userCoOwnerCell;

        tableUsers.setVisible(true);
        //tableUsersSelected.getItems().clear();
        userCol.setCellValueFactory(new PropertyValueFactory<UserTableView, String>("userName"));
        userSelectedCol.setCellValueFactory(new PropertyValueFactory<UserTableView, String>("userName"));

        try {

            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_user_prod.xls")));
            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Case Owner")) {
                    caseOwnerRefCell = i;
                }
                if (filterColName.equals("Co-Owner")){
                    caseCoOwnerRefCell = i;
                }
            }

            ArrayList<String> userArray = new ArrayList<>();

            for (int i = 1; i < lastRow; i++) {

                userCell = filtersheet.getRow(i).getCell(caseOwnerRefCell);
                String userName = userCell.getStringCellValue();
                userCoOwnerCell = filtersheet.getRow(i).getCell(caseCoOwnerRefCell);
                String userCoOwner = "";
                if (userCoOwnerCell != null) {
                    userCoOwner = userCoOwnerCell.getStringCellValue();
                }

                if (!userName.startsWith("PS ") && !userName.startsWith("TS ") && !userName.startsWith("Tech-Ops ")) {
                    userArray.add(userName);
                }
                if (!userCoOwner.equals("")){
                    userArray.add(userCoOwner);
                }
            }

            userArray = (ArrayList) userArray.stream().distinct().collect(Collectors.toList());
            Collections.sort(userArray);

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
                                    tableUsersSelected.getItems().add(selectedUsr);
                                    txtUserSelect.clear();
                                }

                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Unable To Add User to Selected Users By Click...", e);
                            }
                        }
                    }
                });
            });

            btnUsersUpdateAdd.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableUsers.getSelectionModel().getSelectedItem() != null) {
                            UserTableView selectedUsr = tableUsers.getSelectionModel().getSelectedItem();
                            tableUsersSelected.getItems().add(selectedUsr);
                            txtUserSelect.clear();
                        }

                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Add User to Selected Users By Button...", e);
                    }
                }
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
                            logger.log(Level.WARNING, "Unable To Remove User to Selected Users By Click...", e);
                        }
                    }
                }
            });

            btnUsersUpdateRemove.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableUsersSelected.getSelectionModel().getSelectedCells() != null) {
                            UserTableView selectedCust = tableUsersSelected.getSelectionModel().getSelectedItem();
                            tableUsersSelected.getItems().remove(selectedCust);
                        }

                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Remove User to Selected Users By Button...", e);
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
                        logger.log(Level.WARNING, "Unable To Update UserList ...", e);
                    }

                    if (txUsers.getText().equals("")){
                        txUsers.setText(usersFiltered.toString().replace("[", "").replace("]", ""));
                    }else{

                        ArrayList<String> selUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));

                        int selSize = selUser.size();
                        int userFSize = usersFiltered.size();

                        for (int i = 0; i < selSize ; i++) {
                            if (usersFiltered.contains(selUser.get(i))){
                                usersFiltered.remove(selUser.get(i));
                                userFSize--;
                            }
                        }
                        if (userFSize != 0) {
                            txUsers.appendText(", " + usersFiltered.toString().replace("[", "").replace("]", ""));
                        }
                    }

                    pnUsersSelect.setVisible(false);
                    tableUsersSelected.getItems().clear();
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
            logger.log(Level.WARNING, "User Select Array Failed...", e);
        }
    }
    private void workGroupSelect(){
        tableWG.setVisible(true);
        tableWGSelected.setVisible(true);
        WGCol.setCellValueFactory(new PropertyValueFactory<WorkGroupTableView, String>("workGroupName"));
        WGColSelected.setCellValueFactory(new PropertyValueFactory<WorkGroupTableView, String>("workGroupName"));

        HSSFCell wgCell;

        try{
            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(rootFolder + "\\Workgroup\\WorkGroup.xls")));
            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Workgroup Name")) {
                    caseWGRef = i;
                }
                if (filterColName.equals("Vp Current")) {
                    caseVPref = i;
                }

            }

            ArrayList<String> wgArray = new ArrayList<>();

            for (int i = 1; i < lastRow; i++) {

                String vpName = filtersheet.getRow(i).getCell(caseVPref).getStringCellValue();

                if (vpName.equals("Dave Murphy")) {

                    wgCell = filtersheet.getRow(i).getCell(caseWGRef);
                    String workgroupName = "";

                    if (wgCell != null) {
                        workgroupName = wgCell.getStringCellValue();
                    }
                    wgArray.add(workgroupName);
                }
            }

            wgArray = (ArrayList) wgArray.stream().distinct().collect(Collectors.toList());
            Collections.sort(wgArray);

            ObservableList<WorkGroupTableView> wgList = FXCollections.observableArrayList();
            int arraysize = wgArray.size();

            for (int i = 0; i < arraysize; i++) {

                wgList.add(new WorkGroupTableView(wgArray.get(i)));
            }

            FilteredList<WorkGroupTableView> filteredWorkGroups = new FilteredList((ObservableList) wgList, p -> true);
            txtWGSelect.textProperty().addListener((observable, oldValue, newValue) -> {
                filteredWorkGroups.setPredicate(workGroupTableView -> {

                    if (newValue == null || newValue.isEmpty()) {
                        return true;
                    }

                    String lowerCaseCustomerName = newValue.toLowerCase();

                    if (workGroupTableView.getWorkGroupName().toLowerCase().contains(lowerCaseCustomerName)) {
                        return true;
                    }
                    return false;
                });
            });

            SortedList<WorkGroupTableView> sortedWGs = new SortedList<>(filteredWorkGroups);
            sortedWGs.comparatorProperty().bind(tableWG.comparatorProperty());

            tableWG.setItems(filteredWorkGroups);
            tableWG.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
            tableWG.getSelectionModel().setCellSelectionEnabled(true);

            tableWG.getFocusModel().focusedCellProperty().addListener((obs, newVal, oldVal) -> {

                tableWG.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        if (event.getClickCount() > 1) {
                            try {

                                if (tableWG.getSelectionModel().getSelectedItem() != null) {
                                    WorkGroupTableView selectedWG = tableWG.getSelectionModel().getSelectedItem();
                                    //filteredAccounts.add(selectedAcc.getAccountName());
                                    tableWGSelected.getItems().add(selectedWG);
                                }
                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Unable To Add WG to Selected By Click...", e);
                            }
                        }
                    }
                });
            });

            btnWGUpdateAdd.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableWG.getSelectionModel().getSelectedItem() != null) {
                            WorkGroupTableView selectedWG = tableWG.getSelectionModel().getSelectedItem();
                            tableWGSelected.getItems().add(selectedWG);
                            txtWGSelect.clear();
                        }

                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Add WG to Selected By Button...", e);
                    }
                }
            });

            tableWGSelected.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    if (event.getClickCount() > 1) {
                        try {

                            if (tableWGSelected.getSelectionModel().getSelectedCells() != null) {
                                WorkGroupTableView selectedWg = tableWGSelected.getSelectionModel().getSelectedItem();
                                tableWGSelected.getItems().remove(selectedWg);
                            }

                        } catch (Exception e) {
                            logger.log(Level.WARNING, "Unable To Remove WG to Selected By Click...", e);
                        }
                    }
                }
            });

            btnWGUpdateRemove.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableWGSelected.getSelectionModel().getSelectedCells() != null) {
                            WorkGroupTableView selectedWg = tableWGSelected.getSelectionModel().getSelectedItem();
                            tableWGSelected.getItems().remove(selectedWg);
                        }
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Remove WG to Selected By Button...", e);
                    }
                }
            });

            btnWGSelectClear.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    tableWGSelected.getItems().clear();
                }
            });

            btnWGSelectClose.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    pnWGSelect.setVisible(false);
                }
            });

            btnWGUpdate.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    int selected = 0;
                     WgFiltered= new ArrayList<>();

                    try {

                        selected = tableWGSelected.getItems().size();

                        for (int i = 0; i < selected; i++) {

                            WorkGroupTableView addWg = tableWGSelected.getItems().get(i);
                            WgFiltered.add(addWg.getWorkGroupName());

                        }

                        WgFiltered = (ArrayList) WgFiltered.stream().distinct().collect(Collectors.toList());

                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Update Product List...", e);
                    }

                    if(txWorkGroup.getText().equals("")){
                        txWorkGroup.setText(WgFiltered.toString().replace("[", "").replace("]", ""));
                    }else{

                        ArrayList<String> selWg = new ArrayList<>(Arrays.asList(txWorkGroup.getText().split(",\\s*")));

                        int selSize = selWg.size();
                        int wgFSize = WgFiltered.size();

                        for (int i = 0; i < selSize ; i++) {

                            if (WgFiltered.contains(selWg.get(i))){
                                WgFiltered.remove(selWg.get(i));
                                wgFSize--;
                            }
                        }

                        if (wgFSize != 0) {
                            txWorkGroup.appendText(", " + WgFiltered.toString().replace("[", "").replace("]", ""));
                        }
                    }

                    pnWGSelect.setVisible(false);
                    tableWGSelected.getItems().clear();
                    wgToNames();
                }
            });

        }catch (Exception e){
            e.printStackTrace();
        }
    }

    private void wgToNames(){

        HSSFCell wgCell;
        HSSFCell wgUser;

        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(rootFolder + "\\Workgroup\\WorkGroup.xls")));
            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();

            setWG = new ArrayList<>();

            if (!txWorkGroup.getText().isEmpty()) {

                setWG = new ArrayList<>(Arrays.asList(txWorkGroup.getText().split(",\\s*")));
            }

            int userfiltnum = setWG.size();

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Workgroup Name")) {
                    caseWGRef = i;
                }
                if (filterColName.equals("Display Name")) {
                    caseWGNameRef = i;
                }
            }

            setWgNames = new ArrayList<>();

            for (int i = 0; i < userfiltnum ; i++) {

                for (int j = 1; j <lastRow ; j++) {

                    wgCell = filtersheet.getRow(j).getCell(caseWGRef);
                    String work = wgCell.getStringCellValue();

                    if (work.equals(setWG.get(i))) {

                        wgUser = filtersheet.getRow(j).getCell(caseWGNameRef);
                        String user = wgUser.getStringCellValue();
                        setWgNames.add(user);
                    }
                }
            }

            pnWGNames.setVisible(true);
            int wgToNamesArraySize = setWgNames.size();
            txtWGNames.clear();
            for (int i = 0; i <wgToNamesArraySize ; i++) {
                txtWGNames.appendText(setWgNames.get(i)+",");
            }

        }catch (Exception e){
            e.printStackTrace();
        }
    }

    private void accountSelectArray(){

        tableAcc.setVisible(true);
        accCol.setCellValueFactory(new PropertyValueFactory<AccountTableView, String>("accountName"));
        accColSelected.setCellValueFactory(new PropertyValueFactory<AccountTableView, String>("accountName"));

        HSSFCell accCell;

        try {

            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")));
            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();

            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Account Name")) {
                    caseAccountRef = i;
                }
            }

            ArrayList<String> accArray = new ArrayList<>();

            for (int i = 1; i < lastRow; i++) {

                accCell = filtersheet.getRow(i).getCell(caseAccountRef);
                String productName = "";

                if (accCell != null) {
                    productName = accCell.getStringCellValue();
                }
                accArray.add(productName);
            }

            accArray = (ArrayList) accArray.stream().distinct().collect(Collectors.toList());
            Collections.sort(accArray);

            ObservableList<AccountTableView> accList = FXCollections.observableArrayList();

            int arraysize = accArray.size();

            for (int i = 0; i < arraysize; i++) {

                accList.add(new AccountTableView(accArray.get(i)));
            }

            FilteredList<AccountTableView> filteredAccounts = new FilteredList((ObservableList) accList, p -> true);
            txtAccSelect.textProperty().addListener((observable, oldValue, newValue) -> {
                filteredAccounts.setPredicate(accountTableView -> {

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

            SortedList<AccountTableView> sortedProducts = new SortedList<>(filteredAccounts);
            sortedProducts.comparatorProperty().bind(tableAcc.comparatorProperty());

            tableAcc.setItems(filteredAccounts);
            tableAcc.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
            tableAcc.getSelectionModel().setCellSelectionEnabled(true);

            tableAcc.getFocusModel().focusedCellProperty().addListener((obs, newVal, oldVal) -> {

                tableAcc.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        if (event.getClickCount() > 1) {
                            try {

                                if (tableAcc.getSelectionModel().getSelectedItem() != null) {
                                    AccountTableView selectedAccount = tableAcc.getSelectionModel().getSelectedItem();
                                    //filteredAccounts.add(selectedAcc.getAccountName());
                                    tableAccSelected.getItems().add(selectedAccount);
                                }
                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Unable To Add Account to Selected By Click...", e);
                            }
                        }
                    }
                });
            });

            btnAccUpdateAdd.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableAcc.getSelectionModel().getSelectedItem() != null) {
                            AccountTableView selectedAccount = tableAcc.getSelectionModel().getSelectedItem();
                            tableAccSelected.getItems().add(selectedAccount);
                            txtAccSelect.clear();
                        }

                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Add Account to Selected By Button...", e);
                    }
                }
            });

            tableAccSelected.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    if (event.getClickCount() > 1) {
                        try {

                            if (tableAccSelected.getSelectionModel().getSelectedCells() != null) {
                                AccountTableView selectedAcc = tableAccSelected.getSelectionModel().getSelectedItem();
                                tableAccSelected.getItems().remove(selectedAcc);
                            }

                        } catch (Exception e) {
                            logger.log(Level.WARNING, "Unable To Remove Product to Selected By Click...", e);
                        }
                    }
                }
            });

            btnAccUpdateRemove.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableAccSelected.getSelectionModel().getSelectedCells() != null) {
                            AccountTableView selectedCust = tableAccSelected.getSelectionModel().getSelectedItem();
                            tableAccSelected.getItems().remove(selectedCust);
                        }
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Remove Product to Selected By Button...", e);
                    }
                }
            });

            btnAccSelectClear.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    tableAccSelected.getItems().clear();
                }
            });

            btnAccSelectClose.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    pnAccSelect.setVisible(false);
                }
            });

            btnAccUpdate.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    int selected = 0;
                    accountsFiltered = new ArrayList<>();

                    try {

                        selected = tableAccSelected.getItems().size();

                        for (int i = 0; i < selected; i++) {

                            AccountTableView addAcc = tableAccSelected.getItems().get(i);
                            accountsFiltered.add(addAcc.getAccountName());

                        }

                        accountsFiltered = (ArrayList) accountsFiltered.stream().distinct().collect(Collectors.toList());

                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Update Product List...", e);
                    }

                    if(txAccounts.getText().equals("")){
                        txAccounts.setText(accountsFiltered.toString().replace("[", "").replace("]", ""));
                    }else{

                        ArrayList<String> selAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));

                        int selSize = selAcc.size();
                        int accFSize = accountsFiltered.size();

                        for (int i = 0; i < selSize ; i++) {

                            if (accountsFiltered.contains(selAcc.get(i))){
                                accountsFiltered.remove(selAcc.get(i));
                                accFSize--;
                            }
                        }

                        if (accFSize != 0) {
                            txAccounts.appendText(", " + accountsFiltered.toString().replace("[", "").replace("]", ""));
                        }
                    }

                    pnAccSelect.setVisible(false);
                    tableAccSelected.getItems().clear();
                }
            });


        }catch (Exception e){
            e.printStackTrace();
        }
    }

    private void productSelectArray() {

        HSSFCell prodCell;

        tableProducts.setVisible(true);
        //tableProductsSelected.getItems().clear();
        productCol.setCellValueFactory(new PropertyValueFactory<ProductTableView, String>("productName"));
        productColSelected.setCellValueFactory(new PropertyValueFactory<ProductTableView, String>("productName"));

        try {

            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_user_prod.xls")));
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
                String productName = "";

                if (prodCell != null) {
                    productName = prodCell.getStringCellValue();
                }
                prodArray.add(productName);
            }

            prodArray = (ArrayList) prodArray.stream().distinct().collect(Collectors.toList());
            Collections.sort(prodArray);

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
                                    txtProductSelect.clear();
                                }
                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Unable To Add Product to Selected By Click...", e);
                            }
                        }
                    }
                });
            });

            btnProductUpdateAdd.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableProducts.getSelectionModel().getSelectedItem() != null) {
                            ProductTableView selectedProduct = tableProducts.getSelectionModel().getSelectedItem();
                            tableProductsSelected.getItems().add(selectedProduct);
                            txtProductSelect.clear();
                        }

                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Add Product to Selected By Button...", e);
                    }
                }
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
                            logger.log(Level.WARNING, "Unable To Remove Product to Selected By Click...", e);
                        }
                    }
                }
            });

            btnProductUpdateRemove.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableProductsSelected.getSelectionModel().getSelectedCells() != null) {
                            ProductTableView selectedCust = tableProductsSelected.getSelectionModel().getSelectedItem();
                            tableProductsSelected.getItems().remove(selectedCust);
                        }
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Remove Product to Selected By Button...", e);
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
                        logger.log(Level.WARNING, "Unable To Update Product List...", e);
                    }

                    if(txProducts.getText().equals("")){
                        txProducts.setText(productsFiltered.toString().replace("[", "").replace("]", ""));
                    }else{

                        ArrayList<String> selProd = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));

                        int selSize = selProd.size();
                        int productFSize = productsFiltered.size();

                        for (int i = 0; i < selSize ; i++) {

                            if (productsFiltered.contains(selProd.get(i))){
                                productsFiltered.remove(selProd.get(i));
                                productFSize--;
                            }
                        }

                        if (productFSize != 0) {
                            txProducts.appendText(", " + productsFiltered.toString().replace("[", "").replace("]", ""));
                        }
                    }

                    pnProductSelect.setVisible(false);
                    tableProductsSelected.getItems().clear();
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
            logger.log(Level.WARNING, "Product Array Create Failed...", e);
        }
    }

    public void queueSelectArray() {

        tableQueue.setVisible(true);
        //tableQueueSelected.getItems().clear();
        queueCol.setCellValueFactory(new PropertyValueFactory<QueueTableView, String>("queueName"));
        queueColSelected.setCellValueFactory(new PropertyValueFactory<QueueTableView, String>("queueName"));

        int arraySize = queueArray.size();
        Collections.sort(queueArray);

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
                                tableQueueSelected.getItems().add(selectedQue);
                                txtQueueSelect.clear();
                            }
                        } catch (Exception e) {
                            logger.log(Level.WARNING, "Unable To Add Queue to Selected By Click...", e);
                        }
                    }
                }
            });
        });

        btnQueueUpdateAdd.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                try {

                    if (tableQueue.getSelectionModel().getSelectedItem() != null) {
                        QueueTableView selectedQue = tableQueue.getSelectionModel().getSelectedItem();
                        tableQueueSelected.getItems().add(selectedQue);
                        txtQueueSelect.clear();
                    }
                } catch (Exception e) {
                    logger.log(Level.WARNING, "Unable To Add User to Selected By Button...", e);
                }
            }
        });

        tableQueueSelected.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                if (event.getClickCount() > 1){
                    try {

                        if (tableQueueSelected.getSelectionModel().getSelectedCells() != null) {
                            QueueTableView selectedQueue = tableQueueSelected.getSelectionModel().getSelectedItem();
                            tableQueueSelected.getItems().remove(selectedQueue);
                        }
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Remove User to Selected By Click...", e);
                    }
                }
            }
        });

        btnQueueUpdateRemove.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                try {

                    if (tableQueueSelected.getSelectionModel().getSelectedCells() != null) {
                        QueueTableView selectedQueue = tableQueueSelected.getSelectionModel().getSelectedItem();
                        tableQueueSelected.getItems().remove(selectedQueue);
                    }
                } catch (Exception e) {
                    logger.log(Level.WARNING, "Unable To Remove User to Selected By Button...", e);
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
                    logger.log(Level.WARNING, "Unable To Update Queue List...", e);
                }
                if(txQueues.getText().equals("")){
                    txQueues.appendText(queuesFiltered.toString().replace("[", "").replace("]", ""));
                }else{

                    ArrayList<String> selQueue = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));

                    int selSize = selQueue.size();
                    int queueFSize = queuesFiltered.size();

                    for (int i = 0; i < selSize ; i++) {
                        if (queuesFiltered.contains(selQueue.get(i))){
                            queuesFiltered.remove(selQueue.get(i));
                            queueFSize--;
                        }
                    }
                    if (queueFSize != 0) {
                        txQueues.appendText(", " + queuesFiltered.toString().replace("[", "").replace("]", ""));
                    }
                }

                pnQueueSelect.setVisible(false);
                tableQueueSelected.getItems().clear();
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

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

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
                                logger.log(Level.WARNING, "Unable To Add Account to Selected By Click...", e);
                            }
                        }
                    }
                });
            });

            btnFilterAccountUpdateAdd.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableAccounts.getSelectionModel().getSelectedItem() != null) {
                            AccountTableView selectedAcc = tableAccounts.getSelectionModel().getSelectedItem();
                            tableAccountsSelected.getItems().add(selectedAcc);
                        }
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Add Account to Selected By Button...", e);
                    }
                }
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
                            logger.log(Level.WARNING, "Unable To Remove Account to Selected By Click...", e);
                        }
                    }
                }
            });

            btnFilterAccountUpdateRemove.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableAccountsSelected.getSelectionModel().getSelectedCells() != null) {
                            AccountTableView selectedCust = tableAccountsSelected.getSelectionModel().getSelectedItem();
                            tableAccountsSelected.getItems().remove(selectedCust);
                        }
                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Unable To Add Account to Selected By Button...", e);
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
                        logger.log(Level.WARNING, "Unable To Update Account List...", e);
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
            logger.log(Level.WARNING, "Unable To Build Account Array...", e);
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
            pnQueuesSave.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnAccSelect.setVisible(false);
            pnWGSelect.setVisible(false);
            userSelectArray();
            txtUserSelect.requestFocus();
        }

        if (event.getSource() == apnSettings) {
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnQueueSelect.setVisible(false);
            pnQueuesSave.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnProductsSave.setVisible(false);
            pnUsersSave.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnAccSelect.setVisible(false);
            pnAccountLoad.setVisible(false);
            pnAccountSave.setVisible(false);
            pnWGSelect.setVisible(false);
        }
        if (event.getSource() == txProducts) {
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(true);
            pnQueueSelect.setVisible(false);
            pnQueuesSave.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnAccSelect.setVisible(false);
            pnWGSelect.setVisible(false);
            productSelectArray();
            txtProductSelect.requestFocus();
        }

        if (event.getSource() == txQueues) {
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnQueueSelect.setVisible(true);
            pnQueuesSave.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnAccSelect.setVisible(false);
            pnWGSelect.setVisible(false);
            queueSelectArray();
            txtQueueSelect.requestFocus();
        }

        if (event.getSource() == txAccounts){
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnQueueSelect.setVisible(false);
            pnAccSelect.setVisible(true);
            pnQueuesSave.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnWGSelect.setVisible(false);
            accountSelectArray();
            txtAccSelect.requestFocus();
        }
        if (event.getSource() == txWorkGroup){
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnQueueSelect.setVisible(false);
            pnAccSelect.setVisible(false);
            pnQueuesSave.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnWGSelect.setVisible(true);
            workGroupSelect();
            txtWGSelect.requestFocus();
        }
        if (event.getSource() == btnUsersClear) {
            pnUsersSelect.setVisible(false);
            txUsers.clear();
        }
        if (event.getSource() == btnProductsClear) {
            pnProductSelect.setVisible(false);
            txProducts.clear();
        }
        if (event.getSource() == btnQueueClear) {
            pnQueueSelect.setVisible(false);
            txQueues.clear();
        }
        if (event.getSource() == btnWGClear) {
            pnWGSelect.setVisible(false);
            txWorkGroup.clear();
            txtWGNames.clear();
            pnWGNames.setVisible(false);
            pnWGNames.toBack();
        }

        if (event.getSource() == btnAccClear){
            pnAccSelect.setVisible(false);
            txAccounts.clear();
        }

        if (event.getSource() == btnAccountClear) {

            btnCustomerLoad.setVisible(false);
            filteredAccounts.clear();
            tableAccountsSelected.getItems().clear();
            customerText.setText("");
            tableCustomers.setVisible(false);
            btnToExcel.setVisible(false);
            initCustomerNumbers();
            pnAccountSelect.setVisible(false);
        }
        if (event.getSource() == btnClearAll) {
            txUsers.clear();
            txProducts.clear();
            txQueues.clear();
            txAccounts.clear();
            txWorkGroup.clear();
            txtWGNames.setText("");
            regChoice.setValue("NotSet");
        }
        if (event.getSource() == btnSaveDefault) {
            String userFilter = txUsers.getText();
            String queueFilter = txQueues.getText();
            String productFilter = txProducts.getText();
            String accountFilter = txAccounts.getText();
            String regionFilter = regChoice.getValue();
            String workGroupFilter = txtWGNames.getText();

            writeDefaultSettingsToFile(userFilter, queueFilter, productFilter,regionFilter, accountFilter, workGroupFilter);
        }

        if (event.getSource() == btnLoadDefault) {

            readDefaultSettingFiles();
        }

        if(event.getSource() == btnInfo){

            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            ((Stage) alert.getDialogPane().getScene().getWindow()).getIcons().add(new Image("home/image/rbbicon.png"));
            alert.setTitle("RBBN CMT");
            alert.setHeaderText(null);
            alert.setContentText("For any issues/requests please inform us:" + "\n" + "\n" +
                    "Alper Simsek"+ "    " + "asimsek@rbbn.com" + "\n" + "\n" +
                    "Vehbi Benli" + "       " + "vbenli@rbbn.com" + "\n" + "\n" +"RBBN RSD Version 2.0.2");
            alert.showAndWait();
        }

        if (event.getSource() == btnUsersSaveAs){
            pnProductsSave.setVisible(false);
            pnQueuesSave.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnQueueSelect.setVisible(false);
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnUsersSave.toFront();
            pnUsersSave.setVisible(true);
            pnAccountLoad.setVisible(false);
            txtUsersSave.clear();
            saveUserProfile();

        }
        if (event.getSource() == btnUsersSaveClose){
            pnUsersSave.toBack();
            pnUsersSave.setVisible(false);
        }
        if (event.getSource() == btnProductsSaveAs){
            pnProductsSave.toFront();
            pnProductsSave.setVisible(true);
            pnUsersSave.setVisible(false);
            pnQueuesSave.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnQueueSelect.setVisible(false);
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnAccountLoad.setVisible(false);
            txtProductsSave.clear();
            saveProductProfile();

        }
        if (event.getSource() == btnProductsSaveClose){
            pnProductsSave.setVisible(false);
            pnProductsSave.toBack();
        }
        if (event.getSource() == btnQueuesSaveAs){
            pnQueuesSave.toFront();
            pnQueuesSave.setVisible(true);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnQueueSelect.setVisible(false);
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnAccountLoad.setVisible(false);
            txtQueuesSave.clear();
            saveQueueProfile();
        }
        if (event.getSource() == btnQueuesSaveClose){
            pnQueuesSave.setVisible(false);
            pnQueuesSave.toBack();
        }

        if (event.getSource() == btnAccSaveAs){
            pnAccountSave.toFront();
            pnAccountSave.setVisible(true);
            pnQueuesSave.setVisible(false);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnQueueSelect.setVisible(false);
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnAccountLoad.setVisible(false);
            txtQueuesSave.clear();
            saveAccountProfile();
        }
        if (event.getSource() == btnAccountSaveClose){
            pnAccountSave.setVisible(false);
            pnAccountSave.toBack();
        }
        if (event.getSource() == btnUsersLoad){

            pnUsersLoad.toFront();
            pnUsersLoad.setVisible(true);
            pnQueueSelect.setVisible(false);
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnQueuesSave.setVisible(false);
            pnAccountLoad.setVisible(false);
            loadUserProfile();
        }
        if (event.getSource() == btnUsersLoadClose){
            pnUsersLoad.toBack();
            pnUsersLoad.setVisible(false);
        }
        if (event.getSource() == btnProductsLoad){
            pnProductsLoad.toFront();
            pnProductsLoad.setVisible(true);
            pnUsersLoad.setVisible(false);
            pnQueuesLoad.setVisible(false);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnQueuesSave.setVisible(false);
            pnQueueSelect.setVisible(false);
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnAccountLoad.setVisible(false);
            loadProductProfile();

        }
        if (event.getSource() == btnProductsLoadClose){
            pnProductsLoad.toBack();
            pnProductsLoad.setVisible(false);
        }

        if (event.getSource() == btnQueuesLoad){
            pnQueuesLoad.toFront();
            pnQueuesLoad.setVisible(true);
            pnUsersLoad.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnQueuesSave.setVisible(false);
            pnQueueSelect.setVisible(false);
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            pnAccountLoad.setVisible(false);
            loadQueueProfile();

        }
        if (event.getSource() == btnAccLoad){
            pnAccountLoad.toFront();
            pnAccountLoad.setVisible(true);
            pnQueuesLoad.setVisible(false);
            pnUsersLoad.setVisible(false);
            pnProductsLoad.setVisible(false);
            pnUsersSave.setVisible(false);
            pnProductsSave.setVisible(false);
            pnQueuesSave.setVisible(false);
            pnQueueSelect.setVisible(false);
            pnUsersSelect.setVisible(false);
            pnProductSelect.setVisible(false);
            loadAccountProfile();

        }
        if (event.getSource() == btnQueueLoadClose){
            pnQueuesLoad.toBack();
            pnQueuesLoad.setVisible(false);
        }
        if (event.getSource() == btnAccountLoadClose){
            pnAccountLoad.toBack();
            pnAccountLoad.setVisible(false);
        }

        if(event.getSource() == btnManClose){
            apnManLogin.toBack();
        }
        if (event.getSource() == btnManLogin ){

            checkManUser();
        }
        if (event.getSource() == txtpass){
            txtpass.clear();
        }
        if (event.getSource() == btnUnlock && btnUnlock.getGlyphName().equals("LOCK")){
            apnManLogin.toFront();
            txtpass.requestFocus();

            txtpass.setOnKeyPressed(new EventHandler<KeyEvent>() {
                @Override
                public void handle(KeyEvent event) {
                    if (event.getCode() == KeyCode.ENTER){
                        checkManUser();
                    }
                }
            });
        }
    }

    private void checkManUser(){

        String password = "654123";

        if (!txtpass.getText().equals("")){

            String promptedpass = txtpass.getText();
            if (promptedpass.equals(password)){
                apnManLogin.toBack();
                //btnProjection.setVisible(true);
                btnSkillSet.setVisible(true);
                btnUnlock.setGlyphName("UNLOCK");
            }
            else{
            }
        }
    }

    private void loadAccountProfile(){
        accountProfileList.getItems().clear();
        ObservableList<String> accountProfiles = FXCollections.observableArrayList();

        ArrayList<File> files = new ArrayList<>();
        File repo = new File(rootFolder + "\\Profile\\Account");
        File[] fileList = repo.listFiles();

        for (int i = 0 ; i < fileList.length ; i++) {
            accountProfiles.addAll(fileList[i].getName());
        }

        accountProfileList.getItems().addAll(accountProfiles);
        accountProfileList.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
        accountProfileList.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                String selectedProfile = accountProfileList.getSelectionModel().getSelectedItem().toString();

                btnAccountProfLoad.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        File accountProfileFile = new File(rootFolder + "\\Profile\\Account\\" + selectedProfile);

                        txAccounts.clear();

                        if (accountProfileFile.isFile()) {
                            Scanner s = null;
                            try {
                                s = new Scanner(accountProfileFile);
                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/Account - No File Found...: ", e);
                            }
                            while (s.hasNextLine()) {
                                txAccounts.appendText(s.nextLine() + ",");
                            }
                            s.close();
                            pnAccountLoad.toBack();
                            pnAccountLoad.setVisible(false);
                        }
                    }
                });

                btnAccountProfDelete.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        File accountFile = new File(rootFolder + "\\Profile\\Account\\" + selectedProfile);
                        accountFile.delete();
                        accountProfileList.getItems().remove(selectedProfile);
                    }
                });
            }});
    }

    private void loadQueueProfile(){

        queueProfileList.getItems().clear();
        ObservableList<String> queueProfiles = FXCollections.observableArrayList();

        ArrayList<File> files = new ArrayList<File>();
        File repo = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Queue");
        File[] fileList = repo.listFiles();

        for (int i = 0 ; i < fileList.length ; i++) {
            queueProfiles.addAll(fileList[i].getName());
        }

        queueProfileList.getItems().addAll(queueProfiles);
        queueProfileList.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
        queueProfileList.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                String selectedProfile = queueProfileList.getSelectionModel().getSelectedItem().toString();

                btnQueueProfLoad.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        File queueProfileFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Queue\\" + selectedProfile);

                        txQueues.clear();

                        if (queueProfileFile.isFile()) {
                            Scanner s = null;
                            try {
                                s = new Scanner(queueProfileFile);
                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/Queues - No File Found...: ", e);
                            }
                            while (s.hasNextLine()) {
                                txQueues.appendText(s.nextLine() + ",");
                            }
                            s.close();
                            pnQueuesLoad.toBack();
                            pnQueuesLoad.setVisible(false);
                        }
                    }
                });

                btnQueueProfDelete.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        File caseNoteFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Queue\\" + selectedProfile);
                        caseNoteFile.delete();
                        queueProfileList.getItems().remove(selectedProfile);
                    }
                });
            }});
    }

    private void loadProductProfile(){

        productProfileList.getItems().clear();
        ObservableList<String> productProfiles = FXCollections.observableArrayList();

        ArrayList<File> files = new ArrayList<File>();
        File repo = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Product");
        File[] fileList = repo.listFiles();

        for (int i = 0 ; i < fileList.length ; i++) {
            productProfiles.addAll(fileList[i].getName());
        }

        productProfileList.getItems().addAll(productProfiles);
        productProfileList.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
        productProfileList.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                String selectedProfile = productProfileList.getSelectionModel().getSelectedItem().toString();

                btnProdProfLoad.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        File productProfileFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Product\\" + selectedProfile);

                        txProducts.clear();

                        if (productProfileFile.isFile()) {
                            Scanner s = null;
                            try {
                                s = new Scanner(productProfileFile);
                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/Product - No File Found...: ", e);
                            }
                            while (s.hasNextLine()) {
                                txProducts.appendText(s.nextLine() + ",");
                            }
                            s.close();
                            pnProductsLoad.toBack();
                            pnProductsLoad.setVisible(false);
                        }
                    }
                });

                btnProductProfDelete.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {
                        File caseNoteFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Product\\" + selectedProfile);
                        caseNoteFile.delete();
                        productProfileList.getItems().remove(selectedProfile);
                    }
                });
            }});
    }

    private void loadUserProfile(){

        userProfileList.getItems().clear();
        ObservableList<String> userProfiles = FXCollections.observableArrayList();

        //ArrayList<String> profileUsers = new ArrayList<>();

        ArrayList<File> files = new ArrayList<File>();
        File repo = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\User");
        File[] fileList = repo.listFiles();

        if(fileList.length != 0){

            for (int i = 0 ; i < fileList.length ; i++) {
                userProfiles.addAll(fileList[i].getName());
            }

            userProfileList.getItems().addAll(userProfiles);
            userProfileList.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
            userProfileList.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    String selectedProfile = userProfileList.getSelectionModel().getSelectedItem().toString();

                    btnUserProfLoad.setOnMouseClicked(new EventHandler<MouseEvent>() {
                        @Override
                        public void handle(MouseEvent event) {

                            File userProfileFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\User\\" + selectedProfile);

                            txUsers.clear();

                            if (userProfileFile.isFile()) {
                                Scanner s = null;
                                try {
                                    s = new Scanner(userProfileFile);
                                } catch (Exception e) {
                                    logger.log(Level.WARNING, "Profile/User - No File Found...: ", e);
                                }
                                while (s.hasNextLine()) {
                                    txUsers.appendText(s.nextLine() + ",");
                                }
                                s.close();
                                pnUsersLoad.toBack();
                                pnUsersLoad.setVisible(false);
                            }

                        }
                    });

                    btnUserProfDelete.setOnMouseClicked(new EventHandler<MouseEvent>() {
                        @Override
                        public void handle(MouseEvent event) {
                            File caseNoteFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\User\\" + selectedProfile);
                            caseNoteFile.delete();
                            userProfileList.getItems().remove(selectedProfile);
                        }
                    });
                }
            });
        }else{
            alertUser(strLoadProf);
            pnUsersLoad.setVisible(false);
        }

    }

    private void saveAccountProfile(){

        ArrayList<String> setAcc = new ArrayList<>(Arrays.asList(txAccounts.getText().split(",\\s*")));

        btnAccountSave.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                if (!txtAccountSave.getText().isEmpty()) {

                    try {

                        File accProfFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Account\\" + txtAccountSave.getText());

                        if (!accProfFile.exists()) {
                            try {

                                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile").mkdir();
                                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Account").mkdir();
                                FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Account\\" + txtAccountSave.getText()));

                                int size = setAcc.size();

                                for (int i = 0; i < size; i++) {

                                    writer.write(setAcc.get(i) + "\n");

                                }
                                writer.close();

                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/Account - No Folder...: ", e);
                            }
                        } else {
                            try {

                                FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Account\\" + txtAccountSave.getText()));
                                int size = setAcc.size();
                                for (int i = 0; i < size; i++) {

                                    writer.write(setAcc.get(i) + "\n");
                                }
                                writer.close();

                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/Account - Not Able To Save...: ", e);
                            }
                        }

                        pnAccountSave.toBack();
                        pnAccountSave.setVisible(false);

                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                } else{
                    alertUser(strSave);
                }
            }
        });
    }

    private void saveQueueProfile(){

        ArrayList<String> setQue = new ArrayList<>(Arrays.asList(txQueues.getText().split(",\\s*")));

        btnQueuesSave.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                if (!txtQueuesSave.getText().isEmpty()) {

                    try {

                        File userProfFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Queue\\" + txtQueuesSave.getText());

                        if (!userProfFile.exists()) {
                            try {

                                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile").mkdir();
                                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Queue").mkdir();
                                FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Queue\\" + txtQueuesSave.getText()));

                                int size = setQue.size();

                                for (int i = 0; i < size; i++) {

                                    writer.write(setQue.get(i) + "\n");

                                }
                                writer.close();

                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/Queues - No Folder...: ", e);
                            }
                        } else {
                            try {

                                FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Queue\\" + txtQueuesSave.getText()));
                                int size = setQue.size();
                                for (int i = 0; i < size; i++) {

                                    writer.write(setQue.get(i) + "\n");
                                }
                                writer.close();

                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/Queues - Not Able To Save...: ", e);
                            }
                        }

                        pnQueuesSave.toBack();
                        pnQueuesSave.setVisible(false);

                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                } else{
                    alertUser(strSave);
                }
            }
        });
    }
    private void saveProductProfile(){

        ArrayList<String> setProd = new ArrayList<>(Arrays.asList(txProducts.getText().split(",\\s*")));

        btnProductsSave.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                if (!txtProductsSave.getText().isEmpty()) {

                    try {

                        File userProfFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Product\\" + txtProductsSave.getText());

                        if (!userProfFile.exists()) {
                            try {

                                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile").mkdir();
                                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Product").mkdir();
                                FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Product\\" + txtProductsSave.getText()));

                                int size = setProd.size();

                                for (int i = 0; i < size; i++) {

                                    writer.write(setProd.get(i) + "\n");

                                }
                                writer.close();

                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/Product - No Folder...: ", e);
                            }
                        } else {
                            try {

                                FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\Product\\" + txtProductsSave.getText()));
                                int size = setProd.size();
                                for (int i = 0; i < size; i++) {

                                    writer.write(setProd.get(i) + "\n");
                                }
                                writer.close();

                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/Product - Not Able To Save...: ", e);
                            }
                        }
                        pnProductsSave.toBack();
                        pnProductsSave.setVisible(false);

                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Profile/Product - No File...: ", e);
                    }
                }else {
                    alertUser(strSave);
                }
            }
        });
    }

    private void saveUserProfile(){

        ArrayList<String> setUser = new ArrayList<>(Arrays.asList(txUsers.getText().split(",\\s*")));

        btnUsersSave.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {

                if (!txtUsersSave.getText().isEmpty()) {

                    try {

                        File userProfFile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\User\\" + txtUsersSave.getText());

                        if (!userProfFile.exists()) {
                            try {

                                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile").mkdir();
                                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\User").mkdir();
                                FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\User\\" + txtUsersSave.getText()));

                                int size = setUser.size();

                                for (int i = 0; i < size; i++) {

                                    writer.write(setUser.get(i) + "\n");

                                }
                                writer.close();

                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/User - No Folder...: ", e);
                            }
                        } else {
                            try {

                                FileWriter writer = new FileWriter(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Profile\\User\\" + txtUsersSave.getText()));
                                int size = setUser.size();
                                for (int i = 0; i < size; i++) {

                                    writer.write(setUser.get(i) + "\n");
                                }
                                writer.close();

                            } catch (Exception e) {
                                logger.log(Level.WARNING, "Profile/User - Not Able To Save...: ", e);
                            }
                        }

                        pnUsersSave.toBack();
                        pnUsersSave.setVisible(false);

                    } catch (Exception e) {
                        logger.log(Level.WARNING, "Profile/User - No File...: ", e);
                    }
                }
                else{
                    alertUser(strSave);
                }
            }
        });
    }

    private void alertUser(String str){

        Alert alert = new Alert(Alert.AlertType.WARNING);
        ((Stage) alert.getDialogPane().getScene().getWindow()).getIcons().add(new Image("home/image/rbbicon.png"));
        alert.setTitle("RBBN CMT WARNING:");
        alert.setHeaderText(null);
        alert.setContentText(str);
        alert.showAndWait();
    }

    public void setqueueArray() {

        queueArray = new ArrayList<>();

        List queues = Arrays.asList("Kandy NOC", "KBS Onboarding", "KBS Operations","KBS Support", "PS A2 Call Processing", "PS A2 Gateways", "PS A2 GENCOM", "PS A2 IMM",
                "PS A2 OAM", "PS A2 WAM", "PS A6", "PS Billing", "PS C3","PS CBM SDM","PS CCA SST SAM21 Platform", "PS CICM","PS CM9520","PS Converged Intelligent Messaging (CIM)",
                "PS CoreBase SW", "PS CoreHardware", "PS CPaaS", "PS CSLAN8600","PS DMS SS7", "PS DSI NFS","PS EMS","PS EMT", "PS G5", "PS Gateways", "PS GENiUS", "PS GENView Analytics", "PS GPU",
                "PS GSX", "PS GVBM", "PS GVPP","PS GWC", "PS hiG Gateways","PS IN", "PS Intelligent Edge", "PS Kandy","PS Kandy Wrappers", "PS LI / TOPS","PS Lines Services", "PS MG15K G2 G6","PS MG9K",
                "PS MRFP", "PS NSP","PS OAM IEMS","PS OAM SESM","PS OAM SPFS", "PS Protect Netscore","PS PSX","PS Ribbon Protect","PS RSM","PS SBC","PS SeGW", "PS SGX","PS Signaling", "PS SIP Lines/SIP PBX",
                "PS SPiDR CallP","PS SPiDR OAM","PS SPM MG4K","PS SST","PS Trunking","PS UT-SD","PS XLA","PS XPM V52",
                "Tech-Ops ER Support","TS Asia","TS CALA","TS Converged Intelligent Messaging (CIM)","TS EDGE","TS EMEA","TS EMEA Marquee","TS EMEA PI","TS GTAC SERVICES","TS Intelligent Edge","TS Japan Marquee",
                "TS MEXICO","TS MNOC","TS NA","TS NA C15","TS NA DCO","TS NA Federal","TS NA G-Series","TS NA GTD5-5ESS","TS NA Marquee","TS NA Safari","TS NA Safari(GPS)","TS NA S-Series","TS NA Verizon Wireless",
                "TS Non Technical","TS NSP","TS PSD","TS TAC-RESPONSE","TS TAQUA","TS UT-SD");

        queueArray.addAll(queues);
        Collections.sort(queueArray);
    }


    public void setButtons(){
        btnProjection.setVisible(true);
        btnSkillSet.setVisible(true);
    }

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
            engSkillMyTeam();
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
            engSkillOverAllTeam();
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
            btnSkillsExport.setVisible(false);
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
            btnSkillsExport.setVisible(false);
            skillOverAllTeam();
        }
    }

    private void engSkillMyTeam(){

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
                userSkillRef =0;
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
                                    userSkillRef = i;
                                }
                            }

                            for (int i = 1; i <lastRow ; i++) {

                                cellVal = sheet.getRow(i).getCell(userSkillRef);
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
                            logger.log(Level.WARNING, "SkillSet-EngMyteam Unable To Read Data...: ", e);
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

    private void engSkillOverAllTeam(){

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
                userSkillRef =0;
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
                                    userSkillRef = i;
                                }
                            }

                            for (int i = 1; i <lastRow ; i++) {

                                cellVal = sheet.getRow(i).getCell(userSkillRef);

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
                            logger.log(Level.WARNING, "SkillSet-Engteam Unable To Read Data...: ", e);
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
                btnSkillsExport.setVisible(false);

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
                btnSkillsExport.setVisible(false);

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
                        btnSkillsExport.setVisible(true);

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
                            logger.log(Level.WARNING, "SkillSet-SkillMyTeam Unable To Read Data...: ", e);
                        }

                        btnSkillsExport.setOnMouseClicked(new EventHandler<MouseEvent>() {
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
                                                writer.write("\r"+"\n");
                                        }
                                        writer.close();
                                    }

                                    primaryStage.close();
                                } catch (Exception e) {
                                    logger.log(Level.WARNING, "SkillSet-SkillMyTeam Unable To Export Data...: ", e);
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
                btnSkillsExport.setVisible(false);

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
                        btnSkillsExport.setVisible(true);


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
                            logger.log(Level.WARNING, "SkillSet-SkillTeam Unable To Read Data...: ", e);
                        }

                        btnSkillsExport.setOnMouseClicked(new EventHandler<MouseEvent>() {
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
                                                writer.write("\r"+"\n");
                                        }
                                        writer.close();
                                    }

                                    primaryStage.close();
                                } catch (Exception e) {
                                    logger.log(Level.WARNING, "SkillSet-SkillTeam Unable To Export Data...: ", e);
                                }
                            }
                        });
                    }
                });
            }
        });
    }

    private void readSkills(){

        try {

            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(skillSetFolder + "\\SkillsetProfiles.xls")));
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
            logger.log(Level.WARNING, "SkillSet-Read Skills Failed...: ", e);
        }
    }

    private void readUsers(){

        File usersFile = new File(skillSetFolder + "\\users.txt");

        if (usersFile.isFile()) {

            Scanner s = null;
            try {
                s = new Scanner(usersFile);
            } catch (FileNotFoundException e) {
                logger.log(Level.WARNING, "SkillSet-Read Users Failed...: ", e);
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
            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(skillSetFolder + "\\SkillsetProfiles.xls")));
            HSSFSheet sheet = workbook.getSheetAt(0);

            int cellnum = sheet.getRow(0).getLastCellNum();
            int lastRow = sheet.getLastRowNum();

            readOverAllUsers = new ArrayList<>();
            for (int i = 1; i <cellnum ; i++) {
                String userNameColl = sheet.getRow(0).getCell(i).getStringCellValue();
                readOverAllUsers.add(userNameColl);
            }

        }catch(Exception e){
            logger.log(Level.WARNING, "SkillSet-Read All Users Failed...: ", e);
        }
    }

    private void setLevels(){

        ArrayList<String> level = new ArrayList<>();
        List lev = Arrays.asList("EXPERT", "INTERMEDIATE", "BEGINNER", "NONE");
        level.addAll(lev);
        levels.addAll(level);
    }

    private void webViewShow(){

        WebEngine project = projectWeb.getEngine();

        try {
            project.load(String.valueOf(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Forecast\\All_Products_Forecast_M.html").toURI().toURL()));
        }catch (Exception e){
            e.printStackTrace();
        }

        btnProjectRight.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                try{
                    project.load(String.valueOf(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Forecast\\C3_Forecast_M_C3.html").toURI().toURL()));
                    btnProjectLeft.setVisible(true);
                    btnProjectRight.setVisible(false);
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        });

        btnProjectLeft.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                try {
                    project.load(String.valueOf(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Forecast\\All_Products_Forecast_M.html").toURI().toURL()));
                }catch (Exception e){
                    e.printStackTrace();
                }
                btnProjectLeft.setVisible(false);
                btnProjectRight.setVisible(true);
            }
        });
    }

    private void regionChoice(){

        regChoice.setValue("EMEA");

        HSSFCell regCell;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V3.xls")))) {

            HSSFSheet filtersheet = workbook.getSheetAt(0);
            int lastRow = filtersheet.getLastRowNum();
            int cellnum = filtersheet.getRow(0).getLastCellNum();


            for (int i = 0; i < cellnum; i++) {
                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                if (filterColName.equals("Support Theater")) {
                    caseRegionRef = i;
                }
            }

            ArrayList<String> regionArray = new ArrayList<>();

            for (int i = 1; i < lastRow; i++) {

                regCell = filtersheet.getRow(i).getCell(caseRegionRef);
                String regName = regCell.getStringCellValue();
                regionArray.add(regName);
            }

            regionArray = (ArrayList) regionArray.stream().distinct().collect(Collectors.toList());
            Collections.sort(regionArray);

            int size = regionArray.size();

            for (int i = 0; i < size; i++) {
                regChoice.getItems().add(regionArray.get(i));
            }

        }catch (Exception e){
            e.printStackTrace();
        }
    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {

        folderArrange.arrangeCMTFolder();

        try{

            fh = new FileHandler(logFolder + "\\cmt_log", true);
            fh.setFormatter(new SimpleFormatter());
            fh.setLevel(Level.FINE);
            logger.addHandler(fh);

        }catch (Exception e){
            e.printStackTrace();
        }

        logger.info("Program Started");
        setqueueArray();
        readTimeStamp();
        myProductsPage();
        overviewPage();
        myCasesPage();
        regionChoice();
        readDefaultSettingFiles();
    }
}