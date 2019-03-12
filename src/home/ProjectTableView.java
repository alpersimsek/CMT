package home;

import javafx.beans.property.SimpleStringProperty;

public class ProjectTableView {

    private final SimpleStringProperty prjCaseNo;
    private final SimpleStringProperty prjCaseAccount;
    private final SimpleStringProperty prjCaseProduct;
    private final SimpleStringProperty prjCaseSubject;
    private final SimpleStringProperty prjModDate;
    private final SimpleStringProperty prjCaseStatus;
    private final SimpleStringProperty prjCaseSeverity;
    private final SimpleStringProperty prjCaseNumber;
    private final SimpleStringProperty prjHotR;
    private final SimpleStringProperty prjGateDate;
    private final SimpleStringProperty prjRegion;
    private final SimpleStringProperty prjSiteStatus;

    protected ProjectTableView(String cNo, String cAcc, String cProd, String cSub, String cModDate, String cStat, String cSev,
                               String cNum, String cHotR, String cGateDate, String cReg, String cSite) {


        this.prjCaseNo = new SimpleStringProperty(cNo);
        this.prjCaseAccount = new SimpleStringProperty(cAcc);
        this.prjCaseProduct = new SimpleStringProperty(cProd);
        this.prjCaseSubject = new SimpleStringProperty(cSub);
        this.prjModDate = new SimpleStringProperty(cModDate);
        this.prjCaseStatus = new SimpleStringProperty(cStat);
        this.prjCaseSeverity = new SimpleStringProperty(cSev);
        this.prjCaseNumber = new SimpleStringProperty(cNum);
        this.prjHotR = new SimpleStringProperty(cHotR);
        this.prjGateDate = new SimpleStringProperty(cGateDate);
        this.prjRegion = new SimpleStringProperty(cReg);
        this.prjSiteStatus = new SimpleStringProperty(cSite);
    }

    public String getPrjCaseNo() {
        return prjCaseNo.get();
    }

    public void setPrjCaseNo(String cNo) {
        prjCaseNo.set(cNo);
    }

    public String getPrjCaseAccount() {
        return prjCaseAccount.get();
    }

    public void setPrjCaseAccount(String cAcc) {
        prjCaseAccount.set(cAcc);
    }

    public String getPrjCaseProduct() {
        return prjCaseProduct.get();
    }

    public void setPrjCaseProduct(String cProd) {
        prjCaseProduct.set(cProd);
    }

    public String getPrjCaseSubject() {
        return prjCaseSubject.get();
    }

    public void setPrjCaseSubject(String cSubj) {
        prjCaseSubject.set(cSubj);
    }

    public String getPrjModDate() {
        return prjModDate.get();
    }

    public void setPrjModDate(String cModDate) {
        prjModDate.set(cModDate);
    }

    public String getPrjCaseStatus() {
        return prjCaseStatus.get();
    }

    public void setPrjCaseStatus(String cStat) {
        prjCaseStatus.set(cStat);
    }

    public String getPrjCaseSeverity() {
        return prjCaseSeverity.get();
    }

    public void setPrjCaseSeverity(String cSev) {
        prjCaseSeverity.set(cSev);
    }

    public String getPrjCaseNumber() {
        return prjCaseNumber.get();
    }

    public void setPrjCaseNumber(String cNum) {
        prjCaseNumber.set(cNum);
    }

    public String getPrjHotR() {
        return prjHotR.get();
    }

    public void setPrjHotR(String cHotR) {
        prjHotR.set(cHotR);
    }


    public String getPrjGateDate() {
        return prjGateDate.get();
    }

    public void setPrjGateDate(String cGateDate) {
        prjGateDate.set(cGateDate);
    }

    public String getPrjRegion() {
        return prjRegion.get();
    }

    public void setPrjRegion(String cReg) {
        prjRegion.set(cReg);
    }

    public String getPrjSiteStatus() {
        return prjSiteStatus.get();
    }

    public void setPrjSiteStatus(String cSite) {
        prjSiteStatus.set(cSite);
    }

}