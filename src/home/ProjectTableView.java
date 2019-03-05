package home;

import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleObjectProperty;
import javafx.beans.property.SimpleStringProperty;

import java.time.LocalDate;

public class ProjectTableView {

    private final SimpleStringProperty prjCaseNumber;
    private final SimpleStringProperty prjCaseStatus;
    private final SimpleStringProperty prjCaseSeverity;
    private final SimpleStringProperty prjCaseOwner;
    private final SimpleIntegerProperty prjCaseAge;
    private final SimpleStringProperty prjCaseProduct;
    private final SimpleStringProperty prjCaseHotListLevel;
    private final SimpleStringProperty prjCaseHotListReason;
    private final SimpleStringProperty prjCaseHotListComment;
    private final SimpleStringProperty prjCaseHotListBy;
    private final SimpleStringProperty prjCaseHotListDate;
    private final SimpleStringProperty prjGatingDate;
    private final SimpleStringProperty prjCaseAccount;
    private final SimpleStringProperty prjCaseRegion;
    private final SimpleStringProperty prjCaseEscalatedBy;
    private final SimpleStringProperty prjCaseSubject;
    private final SimpleStringProperty prjCaseSupportType;


    protected ProjectTableView(String cNum, String cStat, String cSev, String cOwn, Integer cAge, String cProd, String cHotL, String cHotR, String cHotC, String cHotB, String cHotDate, String cPrjDate,
                               String cAcc, String cReg, String cEsc, String cSub, String cType) {

        this.prjCaseNumber = new SimpleStringProperty(cNum);
        this.prjCaseStatus = new SimpleStringProperty(cStat);
        this.prjCaseSeverity = new SimpleStringProperty(cSev);
        this.prjCaseOwner = new SimpleStringProperty(cOwn);
        this.prjCaseProduct = new SimpleStringProperty(cProd);
        this.prjCaseHotListLevel = new SimpleStringProperty(cHotL);
        this.prjCaseHotListReason = new SimpleStringProperty(cHotR);
        this.prjCaseHotListComment = new SimpleStringProperty(cHotC);
        this.prjCaseHotListBy = new SimpleStringProperty(cHotB);
        this.prjCaseAge = new SimpleIntegerProperty(cAge);
        this.prjCaseHotListDate = new SimpleStringProperty(cHotDate);
        this.prjGatingDate = new SimpleStringProperty(cPrjDate);
        this.prjCaseAccount = new SimpleStringProperty(cAcc);
        this.prjCaseRegion = new SimpleStringProperty(cReg);
        this.prjCaseEscalatedBy = new SimpleStringProperty(cEsc);
        this.prjCaseSubject = new SimpleStringProperty(cSub);
        this.prjCaseSupportType = new SimpleStringProperty(cType);

    }

    public String getPrjCaseNumber() {
        return prjCaseNumber.get();
    }

    public void setPrjCaseNumber(String cNum) {
        prjCaseNumber.set(cNum);
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

    public String getPrjCaseOwner() {
        return prjCaseOwner.get();
    }

    public void setPrjCaseOwner(String cOwn) {
        prjCaseOwner.set(cOwn);
    }

    public String getPrjCaseProduct() {
        return prjCaseProduct.get();
    }

    public void setPrjCaseProduct(String cProd) {
        prjCaseProduct.set(cProd);
    }

    public String getPrjCaseHotListLevel() {
        return prjCaseHotListLevel.get();
    }

    public void setPrjCaseHotListLevel(String cHotL) {
        prjCaseHotListLevel.set(cHotL);
    }

    public String getPrjCaseHotListReason() {
        return prjCaseHotListReason.get();
    }

    public void setPrjCaseHotListReason(String cHotR) {
        prjCaseHotListReason.set(cHotR);
    }

    public String getPrjCaseHotListComment() {
        return prjCaseHotListComment.get();
    }

    public void setPrjCaseHotListComment(String cHotC) {
        prjCaseHotListComment.set(cHotC);
    }

    public String getPrjCaseHotListBy() {
        return prjCaseHotListBy.get();
    }

    public void setPrjCaseHotListBy(String cHotB) {
        prjCaseHotListBy.set(cHotB);
    }

    public Integer getPrjCaseAge() {
        return prjCaseAge.get();
    }

    public void setPrjCaseAge(Integer cAge) {
        prjCaseAge.set(cAge);
    }

    public String getPrjCaseHotListDate() {
        return prjCaseHotListDate.get();
    }

    public void setPrjCaseHotListDate(String cHotD) {
        prjCaseHotListDate.set(cHotD);
    }

    public String getPrjGatingDate() {
        return prjGatingDate.get();
    }

    public void getPrjGatingDate(String cPrjDate) {
        prjGatingDate.set(cPrjDate);
    }

    public String getPrjCaseAccount() {
        return prjCaseAccount.get();
    }

    public void getPrjCaseAccount(String cAcc) {
        prjCaseAccount.set(cAcc);
    }

    public String getPrjCaseRegion() {
        return prjCaseRegion.get();
    }

    public void setPrjCaseRegion(String cReg) {
        prjCaseRegion.set(cReg);
    }

    public String getPrjCaseEscalatedBy() {
        return prjCaseEscalatedBy.get();
    }

    public void setPrjCaseEscalatedBy(String cEsc) {
        prjCaseEscalatedBy.set(cEsc);
    }

    public String getPrjCaseSubject() {
        return prjCaseSubject.get();
    }

    public void setPrjCaseSubject(String cSubj) {
        prjCaseSubject.set(cSubj);
    }

    public String getPrjCaseSupportType() {
        return prjCaseSupportType.get();
    }

    public void setPrjCaseSupportType(String cType) {
        prjCaseSupportType.set(cType);
    }
}