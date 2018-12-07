package home;

import javafx.beans.property.SimpleStringProperty;

public class UserTableView {

    private final SimpleStringProperty userName;

    protected UserTableView(String usrName) {
        this.userName = new SimpleStringProperty(usrName);
    }

    public String getUserName() {
        return userName.get();
    }
    public void setUserName(String usrName) {
        userName.set(usrName);
    }

}