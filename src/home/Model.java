package home;

import de.jensd.fx.glyphs.fontawesome.FontAwesomeIconView;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.collections.transformation.SortedList;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.AnchorPane;
import javafx.scene.web.WebEngine;
import javafx.scene.web.WebView;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collections;
import java.util.ResourceBundle;
import java.util.stream.Collectors;

public class Model implements Initializable {

    @FXML
    private TextField txtProducts;

    @FXML
    private TextField txtFilter;

    @FXML
    private RadioButton rdAll;

    @FXML
    private Button btnRun;

    @FXML
    private AnchorPane apnFCProdSelect;

    @FXML
    private TableView<ProductTableView> tableForecastProd;

    @FXML
    private TableColumn<ProductTableView, String> tableColumn;

    @FXML
    private Button btnUpdatetxt;

    @FXML
    private FontAwesomeIconView btnaddSelected;

    @FXML
    private FontAwesomeIconView btnremoveSelected;

    @FXML
    private TableView<ProductTableView> tableFCProdSelected;

    @FXML
    private TableColumn<ProductTableView, String> tableFCColumn;

    Boolean allProducts = false;


    @FXML
    private WebView webimage;

    ArrayList<String> productsFCFiltered = new ArrayList<String>();

    private void selectForecastArray(){

        int caseProductRef = 0;
        HSSFCell prodCell;
        txtProducts.setEditable(true);

        tableColumn.setCellValueFactory(new PropertyValueFactory<ProductTableView, String>("productName"));
        tableFCColumn.setCellValueFactory(new PropertyValueFactory<ProductTableView, String>("productName"));

        try {

            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(System.getProperty("user.home") + "\\Documents\\CMT\\Data\\cmt_user_prod.xls")));
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

            Collections.sort(prodArray);

            ObservableList<ProductTableView> prodList = FXCollections.observableArrayList();

            int arraysize = prodArray.size();

            for (int i = 0; i < arraysize; i++) {

                prodList.add(new ProductTableView(prodArray.get(i)));
            }

            FilteredList<ProductTableView> filteredProducts = new FilteredList((ObservableList) prodList, p -> true);
            txtFilter.textProperty().addListener((observable, oldValue, newValue) -> {
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
            sortedProducts.comparatorProperty().bind(tableForecastProd.comparatorProperty());

            tableForecastProd.setItems(filteredProducts);
            tableForecastProd.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
            tableForecastProd.getSelectionModel().setCellSelectionEnabled(true);

            tableForecastProd.getFocusModel().focusedCellProperty().addListener((obs, newVal, oldVal) -> {

                tableForecastProd.setOnMouseClicked(new EventHandler<MouseEvent>() {
                    @Override
                    public void handle(MouseEvent event) {

                        if (event.getClickCount() > 1) {
                            try {

                                if (tableForecastProd.getSelectionModel().getSelectedItem() != null) {
                                    ProductTableView selectedProduct = tableForecastProd.getSelectionModel().getSelectedItem();
                                    //filteredAccounts.add(selectedAcc.getAccountName());
                                    tableFCProdSelected.getItems().add(selectedProduct);
                                }

                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }
                    }
                });

            });

            btnaddSelected.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableForecastProd.getSelectionModel().getSelectedItem() != null) {
                            ProductTableView selectedProduct = tableForecastProd.getSelectionModel().getSelectedItem();
                            tableFCProdSelected.getItems().add(selectedProduct);
                        }

                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            tableFCProdSelected.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    if (event.getClickCount() > 1) {
                        try {

                            if (tableFCProdSelected.getSelectionModel().getSelectedCells() != null) {
                                ProductTableView selectedCust = tableFCProdSelected.getSelectionModel().getSelectedItem();
                                tableFCProdSelected.getItems().remove(selectedCust);
                            }

                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }

                }
            });

            btnremoveSelected.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    try {

                        if (tableFCProdSelected.getSelectionModel().getSelectedCells() != null) {
                            ProductTableView selectedCust = tableFCProdSelected.getSelectionModel().getSelectedItem();
                            tableFCProdSelected.getItems().remove(selectedCust);
                        }

                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });

            btnUpdatetxt.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {

                    int selected = 0;
                    productsFCFiltered = new ArrayList<>();

                    try {

                        selected = tableFCProdSelected.getItems().size();

                        for (int i = 0; i < selected; i++) {

                            ProductTableView addUsr = tableFCProdSelected.getItems().get(i);
                            productsFCFiltered.add(addUsr.getProductName());

                        }

                        productsFCFiltered = (ArrayList) productsFCFiltered.stream().distinct().collect(Collectors.toList());

                    } catch (Exception e) {
                        e.printStackTrace();
                    }

                    txtProducts.setText(productsFCFiltered.toString().replace("[", "").replace("]", ""));
                    apnFCProdSelect.setVisible(false);
                }
            });

            /*btnProductSelectClear.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    tableProductsSelected.getItems().clear();
                }
            });*/

            /*btnProductSelectClose.setOnMouseClicked(new EventHandler<MouseEvent>() {
                @Override
                public void handle(MouseEvent event) {
                    pnProductSelect.setVisible(false);
                }
            });*/

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @FXML
    void handleMouseClicked(MouseEvent event) {

        if (event.getSource() == txtProducts){
            apnFCProdSelect.toFront();
            apnFCProdSelect.setVisible(true);
            txtFilter.requestFocus();
            rdAll.setSelected(false);
            allProducts = false;
            webimage.setVisible(false);
        }
        if (event.getSource() == btnRun){
            runForecast();
        }

    }

    private void runForecast() {

        if (txtProducts.getText().isEmpty() && !rdAll.isSelected()){
            System.out.println("Please Select Product(s) or Select All Product button to proceed");
        }

        if (!txtProducts.getText().isEmpty() && rdAll.isSelected()){
            System.out.println("All Products Selected, ignoring prompted product list");
        }
        if (txtProducts.getText().isEmpty() && rdAll.isSelected()){
            System.out.println("All Products Selected");
        }
        if (!txtProducts.getText().isEmpty() && !rdAll.isSelected()){
            System.out.println(txtProducts.getText());
        }

        webimage.setVisible(true);

        WebEngine webEngine = webimage.getEngine();
        try {
            webEngine.load(String.valueOf(new File("D:\\Users\\asimsek\\PycharmProjects\\rbbnCrystalBall-master\\data\\Forecast_in_Future_M.html").toURI().toURL()));
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    @FXML
    void handleWebClick(MouseEvent event) {

    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {

        selectForecastArray();
    }
}