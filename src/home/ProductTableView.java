package home;

import javafx.beans.property.SimpleStringProperty;

public class ProductTableView {

    private final SimpleStringProperty productName;

    protected ProductTableView(String prdName) {
        this.productName = new SimpleStringProperty(prdName);
    }

    public String getProductName() {
        return productName.get();
    }
    public void setProductName(String prdName) {
        productName.set(prdName);
    }

}