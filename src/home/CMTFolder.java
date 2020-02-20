package home;

import java.io.File;
import java.nio.channels.SelectableChannel;

import static home.Controller.*;

public class CMTFolder {

    public void arrangeCMTFolder() {


        if (Controller.repo.exists()) {

            File repo1 = new File(System.getProperty("user.home") + "\\Documents\\CMT");
            File repo2 = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Settings");
            File repo3 = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Data");
            File repo4 = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Selection");
            File repo5 = new File(System.getProperty("user.home") + "\\Documents\\CMT\\SkilLSet");
            File repo6 = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Log");
            File repo7 = new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes");
            File repo8 = new File(System.getProperty("user.home") + "\\Documents\\CMT\\CaseDetails");


            if (!repo1.exists()) {
                logger.info("CMT Folder Not Valid, creating folder...");
                new File(System.getProperty("user.home") + "\\Documents\\CMT").mkdir();
            }

            if (!repo2.exists()){
                logger.info("Settings Folder Not Valid, creating folder...");
                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Settings").mkdir();
            }

            if (!repo3.exists()){
                logger.info("Data Folder Not Valid, creating folder...");
                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Data").mkdir();
            }

            if (!repo4.exists()){
                logger.info("Selection Folder Not Valid, creating folder...");
                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Selection").mkdir();
            }

            if (!repo5.exists()){
                logger.info("SkillSet Folder Not Valid, creating folder...");
                new File(System.getProperty("user.home") + "\\Documents\\CMT\\SkillSet").mkdir();
            }

            if (!repo6.exists()){
                logger.info("Log Folder Not Valid, creating folder...");
                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Log").mkdir();
            }
            if (!repo7.exists()){
                logger.info("Notes Folder Not Valid, creating folder...");
                new File(System.getProperty("user.home") + "\\Documents\\CMT\\Notes").mkdir();
            }
            if (!repo8.exists()){
                logger.info("Notes Folder Not Valid, creating folder...");
                new File(System.getProperty("user.home") + "\\Documents\\CMT\\CaseDetails").mkdir();
            }

            //Set the folder variables
            Controller.rootFolder = System.getProperty("user.home") + "\\Documents\\CMT";
            Controller.settingsFolder = System.getProperty("user.home") + "\\Documents\\CMT\\Settings";
            Controller.dataFolder = System.getProperty("user.home") + "\\Documents\\CMT\\Data";
            Controller.selectionFolder = System.getProperty("user.home") + "\\Documents\\CMT\\Selection";
            Controller.skillSetFolder = System.getProperty("user.home") + "\\Documents\\CMT\\SkillSet";
            Controller.logFolder = System.getProperty("user.home") + "\\Documents\\CMT\\Log";
            Controller.noteFolder = System.getProperty("user.home") + "\\Documents\\CMT\\Notes";
            Controller.detailsFolder = System.getProperty("user.home") + "\\Documents\\CMT\\CaseDetails";

            File[] fileList = repo1.listFiles();

            for (int i = 0; i < fileList.length; i++) {

                if (fileList[i].getName().equals("cmt_product_default_settings.txt")
                        || fileList[i].getName().equals("cmt_queueu_default_settings.txt") ||
                        fileList[i].getName().equals("cmt_user_default_settings.txt")) {

                    File oldone = new File(System.getProperty("user.home") + "\\Documents\\CMT\\"+fileList[i].getName());
                    oldone.renameTo(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Settings\\"+fileList[i].getName()));
                }
                if (fileList[i].getName().equals("cmt_case_data_V2.csv") || fileList[i].getName().equals("cmt_case_data_V2.xls") || fileList[i].getName().equals("cmt_case_data_V3.xls") ||
                        fileList[i].getName().equals("cmt_comments.csv") || fileList[i].getName().equals("cmt_comments.xls") || fileList[i].getName().equals("cmt_projects.csv") ||
                        fileList[i].getName().equals("cmt_projects.xls")|| fileList[i].getName().equals("cmt_user_prod.csv") || fileList[i].getName().equals("cmt_user_prod.xls")){

                    File oldone2 = new File(System.getProperty("user.home") + "\\Documents\\CMT\\"+fileList[i].getName());
                    oldone2.renameTo(new File(System.getProperty("user.home") + "\\Documents\\CMT\\Data\\"+fileList[i].getName()));
                }
            }

        } else {

            File pro1 = new File(System.getenv("USERPROFILE") + "\\CMT");
            File pro2 = new File(System.getenv("USERPROFILE") + "\\CMT\\Settings");
            File pro3 = new File(System.getenv("USERPROFILE") + "\\CMT\\Data");
            File pro4 = new File(System.getenv("USERPROFILE") + "\\CMT\\Selection");
            File pro5 = new File(System.getenv("USERPROFILE") + "\\CMT\\SkillSet");
            File pro6 = new File(System.getenv("USERPROFILE") + "\\CMT\\Log");
            File pro7 = new File(System.getenv("USERPROFILE") + "\\CMT\\Notes");
            File pro8 = new File(System.getenv("USERPROFILE") + "\\CMT\\CaseDetails");


            if (!pro1.exists()) {
                logger.info("CMT Folder Not Valid, creating folder...");
                new File(System.getenv("USERPROFILE") + "\\CMT").mkdir();
            }
            if (!pro2.exists()){
                logger.info("Settings Folder Not Valid, creating folder...");
                new File(System.getenv("USERPROFILE") + "\\CMT\\Settings").mkdir();
            }

            if (!pro3.exists()){
                logger.info("Data Folder Not Valid, creating folder...");
                new File(System.getenv("USERPROFILE") + "\\CMT\\Settings\\Data").mkdir();
            }

            if (!pro4.exists()){
                logger.info("Selection Folder Not Valid, creating folder...");
                new File(System.getenv("USERPROFILE") + "\\CMT\\Selection").mkdir();
            }

            if (!pro5.exists()){
                logger.info("SkillSet Folder Not Valid, creating folder...");
                new File(System.getenv("USERPROFILE") + "\\CMT\\SkillSet").mkdir();
            }

            if (!pro6.exists()){
                logger.info("Log Folder Not Valid, creating folder...");
                new File(System.getenv("USERPROFILE") + "\\CMT\\SkillSet\\Log").mkdir();
            }
            if (!pro7.exists()){
                logger.info("Log Folder Not Valid, creating folder...");
                new File(System.getenv("USERPROFILE") + "\\CMT\\SkillSet\\Notes").mkdir();
            }
            if (!pro8.exists()){
                logger.info("Log Folder Not Valid, creating folder...");
                new File(System.getenv("USERPROFILE") + "\\CMT\\SkillSet\\CaseDetails").mkdir();
            }
            Controller.rootFolder = System.getenv("USERPROFILE") + "\\CMT";
            Controller.settingsFolder = System.getenv("USERPROFILE") + "\\CMT\\Settings";
            Controller.dataFolder = System.getenv("USERPROFILE") + "\\CMT\\Settings\\Data";
            Controller.selectionFolder = System.getenv("USERPROFILE") + "\\CMT\\Selection";
            Controller.skillSetFolder = System.getenv("USERPROFILE") + "\\CMT\\SkillSet";
            Controller.logFolder = System.getenv("USERPROFILE") + "\\CMT\\SkillSet\\Log";
            Controller.noteFolder = System.getenv("USERPROFILE") + "\\CMT\\SkillSet\\Notes";
            Controller.detailsFolder = System.getenv("USERPROFILE") + "\\CMT\\SkillSet\\CaseDetails";
        }
    }
}
