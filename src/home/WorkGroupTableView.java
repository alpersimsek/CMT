package home;

import javafx.beans.property.SimpleStringProperty;

public class WorkGroupTableView {

    private final SimpleStringProperty workGroupName;

    protected WorkGroupTableView(String wgName) {
        this.workGroupName = new SimpleStringProperty(wgName);
    }

    public String getWorkGroupName() {
        return workGroupName.get();
    }
    public void setWorkGroupName(String wgName) {
        workGroupName.set(wgName);
    }
}
