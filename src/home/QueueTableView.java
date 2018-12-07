package home;

import javafx.beans.property.SimpleStringProperty;

public class QueueTableView {

    private final SimpleStringProperty queueName;

    protected QueueTableView(String queName) {
        this.queueName = new SimpleStringProperty(queName);
    }

    public String getQueueName() {
        return queueName.get();
    }
    public void setQueueName(String queName) {
        queueName.set(queName);
    }

}