package home;

import javafx.beans.property.SimpleStringProperty;

public class AccountTableView {

    private final SimpleStringProperty accountName;

    protected AccountTableView(String accName) {
        this.accountName = new SimpleStringProperty(accName);
    }

    public String getAccountName() {
        return accountName.get();
    }
    public void setAccountName(String accName) {
        accountName.set(accName);
    }

}