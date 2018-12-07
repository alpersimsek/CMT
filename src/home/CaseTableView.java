package home;

import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleObjectProperty;
import javafx.beans.property.SimpleStringProperty;

import java.time.LocalDate;

public class CaseTableView {

    //private final SimpleIntegerProperty caseCount;
    private final SimpleStringProperty caseNumber;
    private final SimpleStringProperty caseStatus;
    private final SimpleStringProperty caseSeverity;
    private final SimpleStringProperty caseResponsible;
    private final SimpleStringProperty caseOwner;
    private final SimpleStringProperty caseEscalatedBy;
    private final SimpleStringProperty caseHotList;
    private final SimpleIntegerProperty caseAge;
    private final SimpleStringProperty caseProduct;
    private final SimpleStringProperty caseAccount;
    private final SimpleStringProperty caseSubject;
    private final SimpleStringProperty caseSupportType;
    private final SimpleObjectProperty nextCaseUpdate;
    private final SimpleStringProperty caseDateTimeOpened;
    private final SimpleStringProperty caseRegion;
    private final SimpleStringProperty caseSecurity;
    private final SimpleStringProperty caseOutFollow;



    protected CaseTableView(String cNum, String cSev, String cStat, String cOwn, String cResp, Integer cAge, LocalDate cNextUp, String cEsc, String cHot, String cOutF, String cSType, String cProd,
                            String cSubj, String cAcc, String cReg, String cSecur, String cOpDat) {
        //this.caseCount = new SimpleIntegerProperty(cCount);
        this.caseNumber = new SimpleStringProperty(cNum);
        this.caseStatus = new SimpleStringProperty(cStat);
        this.caseSeverity = new SimpleStringProperty(cSev);
        this.caseResponsible = new SimpleStringProperty(cResp);
        this.caseOwner = new SimpleStringProperty(cOwn);
        this.caseEscalatedBy = new SimpleStringProperty(cEsc);
        this.caseHotList = new SimpleStringProperty(cHot);
        this.caseAge = new SimpleIntegerProperty(cAge);
        this.caseProduct = new SimpleStringProperty(cProd);
        this.caseAccount = new SimpleStringProperty(cAcc);
        this.caseSubject = new SimpleStringProperty(cSubj);
        this.caseSupportType = new SimpleStringProperty(cSType);
        this.nextCaseUpdate = new SimpleObjectProperty(cNextUp);
        this.caseDateTimeOpened = new SimpleStringProperty(cOpDat);
        this.caseRegion = new SimpleStringProperty(cReg);
        this.caseSecurity = new SimpleStringProperty(cSecur);
        this.caseOutFollow = new SimpleStringProperty(cOutF);
    }

    /*public Integer getCaseCount(){
        return  caseCount.intValue();
    }

    /*public void setCaseCount(Integer cCount){
        caseCount.set(cCount);
    }*/

    public String getCaseNumber() {
        return caseNumber.get();
    }

    public void setCaseNumber(String cNum) {
        caseNumber.set(cNum);
    }
    public String getCaseStatus() {
        return caseStatus.get();
    }

    public void setCaseStatus(String cStat) {
        caseStatus.set(cStat);
    }

    public String getCaseSeverity() {
        return caseSeverity.get();
    }

    public void setCaseSeverity(String cSev) {
        caseSeverity.set(cSev);
    }

    public String getCaseResponsible() {
        return caseResponsible.get();
    }

    public void setCaseResponsible(String cResp) {
        caseResponsible.set(cResp);
    }

    public String getCaseOwner() {
        return caseOwner.get();
    }

    public void setCaseOwner(String cOwn) {
        caseOwner.set(cOwn);
    }

    public String getCaseEscalatedBy() {
        return caseEscalatedBy.get();
    }

    public void setCaseEscalatedBy(String cEsc) {
        caseEscalatedBy.set(cEsc);
    }

    public String getCaseHotList() {
        return caseHotList.get();
    }

    public void setCaseHotList(String cHot) {
        caseHotList.set(cHot);
    }

    public Integer getCaseAge() {
        return caseAge.get();
    }

    public void setCaseAge(Integer cAge) {
        caseAge.set(cAge);
    }

    public String getCaseProduct() {
        return caseProduct.get();
    }

    public void setCaseProduct(String cProd) {
        caseProduct.set(cProd);
    }

    public String getCaseAccount() {
        return caseAccount.get();
    }

    public void setCaseAccount(String cAcc) {
        caseAccount.set(cAcc);
    }

    public String getCaseSubject() {
        return caseSubject.get();
    }

    public void setCaseSubject(String cSubj) {
        caseSubject.set(cSubj);
    }

    public String getCaseSupportType() {
        return caseSupportType.get();
    }

    public void setCaseSupportType(String cSType) {
        caseSupportType.set(cSType);
    }

    public LocalDate getNextCaseUpdate() {
        return (LocalDate) nextCaseUpdate.get();
    }

    public void setNextCaseUpdate(LocalDate cNextUp) {
        nextCaseUpdate.set(cNextUp);
    }

    public String getCaseDateTimeOpened() {
        return caseDateTimeOpened.get();
    }

    public void setCaseDateTimeOpened(String cOpDat) {
        caseDateTimeOpened.set(cOpDat);
    }

    public String getCaseRegion() {
        return caseRegion.get();
    }

    public void setCaseRegion(String cReg) {
        caseRegion.set(cReg);
    }
    public String getCaseSecurity() {
        return caseSecurity.get();
    }

    public void setCaseSecurity(String cSecur) {
        caseSecurity.set(cSecur);
    }

    public String getCaseOutFollow() {
        return caseOutFollow.get();
    }

    public void setCaseOutFollow(String cOutF) {
        caseOutFollow.set(cOutF);
    }

}