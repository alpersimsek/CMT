package home;

import com.univocity.parsers.csv.CsvParser;
import com.univocity.parsers.csv.CsvParserSettings;
import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.util.Duration;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.net.URL;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.logging.Level;
import java.util.regex.Pattern;

import static home.Controller.dataFolder;
import static home.Controller.logger;

public class DownData {

    Timeline time = new Timeline();
    ExecutorService service = Executors.newSingleThreadExecutor();
    int caseNumCellRefData = 0;
    int mycaseAgeRefCellData = 0;


    public void downloadCSV() {

        String filename1 = "cmt_projects.csv";
        String filename2 = "cmt_user_prod.csv";
        String filename3 = "cmt_case_data_V2.csv";
        String filename4 = "cmt_comments.csv";
        String newLoc2 = "https://rbbn.my.salesforce.com/00OC0000006r1xS?export=1&enc=UTF-8&xf=csv?filename=" + filename2;
        String newLoc = "https://rbbn.my.salesforce.com/00OC0000007My3o?export=1&enc=UTF-8&xf=csv?filename=" + filename1;
        String newLoc3 = "https://rbbn.my.salesforce.com/00OC00000076uIg?export=1&enc=UTF-8&xf=csv?filename=" + filename3;
        String newLoc4 = "https://rbbn.my.salesforce.com/00OC0000006r5ig?export=1&enc=UTF-8&xf=csv?filename=" + filename4;

        try {

            FileUtils.copyURLToFile(new URL(newLoc2), new File(dataFolder + "\\cmt_user_prod.csv"));

        } catch (Exception e) {
            logger.log(Level.WARNING, "Could not Download User/Product File", e);
        }

        //Downloaded User Data...Now Parsing...
        logger.info("User Data Download Completed! Now Parsing...");
        parseUserData();

        try {

            FileUtils.copyURLToFile(new URL(newLoc), new File(dataFolder + "\\cmt_projects.csv"));

        } catch (Exception e) {
            logger.log(Level.WARNING, "Could not Download Projects File", e);
        }

        //Downloaded Project Data...Now Parsing...
        logger.info("Project Data Download Completed! Now Parsing...");

        parseProjectData();

        try{

            FileUtils.copyURLToFile(new URL(newLoc3), new File(dataFolder + "\\cmt_case_data_V2.csv"));
            LocalDate refreshDate = LocalDate.now();
            DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm");

            String dataDate = "Data Time Stamp is:" + "\n" + LocalTime.now().format(dtf).toString() + "\n" + refreshDate.toString();

            FileWriter writer = new FileWriter(new File(dataFolder + "\\cmt_data_Date.txt"));
            writer.write(dataDate);
            writer.close();

        }catch (Exception e){
            logger.log(Level.WARNING, "Case Data Download Failed...", e);
        }

        //Downloaded Case Data...Now Parsing...
        logger.info("Case Data Download Completed! Now Parsing...");

        parseData();

        try{

            FileUtils.copyURLToFile(new URL(newLoc4), new File(dataFolder + "\\cmt_comments.csv"));

        }catch (Exception e){
            logger.log(Level.WARNING, "Could not Download Work Notes File", e);
        }

        logger.info("Comment Data Download Completed! Now Parsing...");
        parseComments();
        logger.info("Account Rectify...");
        rectifyAccountNames();

        time = new Timeline();
        time.setCycleCount(Timeline.INDEFINITE);
        time.getKeyFrames().add(new KeyFrame(Duration.minutes(5), new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                time.stop();
                logger.info("Time-Out! Downloading Latest Reports!...");
                service.submit(DownData.this::downloadCSV);
            }
        }));
        time.playFromStart();
    }


    private void parseUserData() {
        try {
            File csvfile = new File(dataFolder + "\\cmt_user_prod.csv");

            HSSFWorkbook workBook = new HSSFWorkbook();

            String xlsFileAddress = dataFolder + "\\cmt_user_prod.xls";
            HSSFSheet sheet = workBook.createSheet("UserProd");

            BufferedReader br = new BufferedReader(new FileReader(csvfile));
            String line;

            int RowNum = 0;

            while ((line = br.readLine()) != null) {
                //line = line.replace("Ã©n", "e");
                line = line.replaceAll("^\"|\"$", "");

                String[] fields = parseCsvLine(line);

                HSSFRow currentRow = sheet.createRow(RowNum);
                for (int i = 0; i < fields.length; i++) {
                    currentRow.createCell(i).setCellValue(fields[i]);
                }
                RowNum++;
            }

            int lastRow = sheet.getLastRowNum();

            for (int i = 0; i < 7; i++) {
                sheet.removeRow(sheet.getRow(lastRow - i));
            }
            FileOutputStream fileOutputStream = new FileOutputStream(xlsFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();

        } catch (Exception e) {
            logger.log(Level.WARNING, "User Data parse failed...", e);
        }
    }

    private void parseComments(){
        try {

            File csvfile = new File(dataFolder + "\\cmt_comments.csv");

            HSSFWorkbook workBook = new HSSFWorkbook();
            String xlsFileAddress = dataFolder + "\\cmt_comments.xls";
            HSSFSheet sheet = workBook.createSheet("Survey");
            CreationHelper helper = workBook.getCreationHelper();

            int r = 0;

            CsvParserSettings settings = new CsvParserSettings();
            settings.setMaxCharsPerColumn(100000);
            settings.getFormat().setLineSeparator("\n");

            CsvParser parser = new CsvParser(settings);
            parser.beginParsing(csvfile);

            String[] row;
            while ((row = parser.parseNext()) != null) {

                Row frow = sheet.createRow((short) r++);
                for (int i = 0; i <row.length ; i++) {
                    frow.createCell(i).setCellValue(helper.createRichTextString(row[i]));
                }
            }

            parser.stopParsing();

            int lastRow = sheet.getLastRowNum();
            for (int i = 0; i < 7; i++) {
                sheet.removeRow(sheet.getRow(lastRow - i));
            }

            FileOutputStream fileOutputStream = new FileOutputStream(xlsFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();

        }catch (Exception e) {
            logger.log(Level.WARNING, "Work Note Parse Failed...", e);
        }
    }

    // Splitting the CSV File for reading
    private String[] parseCsvLine(String line) throws IOException {

        Pattern p = Pattern.compile("\",\"");
        //Pattern p = Pattern.compile(",(?=([^\"]*\"[^\"]*\")*(?![^\"]*\"))");

        // Split input with the pattern
        String[] fields = p.split(line);

        return fields;
    }

    public void parseProjectData(){

        try {

            File csvfile = new File(dataFolder + "\\cmt_projects.csv");
            HSSFWorkbook workBook = new HSSFWorkbook();
            String xlsFileAddress = dataFolder + "\\cmt_projects.xls";
            HSSFSheet sheet = workBook.createSheet("Projects");
            CreationHelper helper = workBook.getCreationHelper();

            int r = 0;

            CsvParserSettings settings = new CsvParserSettings();
            settings.setMaxCharsPerColumn(100000);
            settings.getFormat().setLineSeparator("\n");

            CsvParser parser = new CsvParser(settings);
            parser.beginParsing(csvfile);

            String[] row;
            while ((row = parser.parseNext()) != null) {

                Row frow = sheet.createRow((short) r++);
                for (int i = 0; i <row.length ; i++) {
                    frow.createCell(i).setCellValue(helper.createRichTextString(row[i]));
                }
            }

            parser.stopParsing();

            int lastRow = sheet.getLastRowNum();

            for (int i = 0; i < 5; i++) {
                sheet.removeRow(sheet.getRow(lastRow - i));
            }

            FileOutputStream fileOutputStream = new FileOutputStream(xlsFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();

        }catch (Exception e) {
            logger.log(Level.WARNING, "Project Data parse failed...", e);
        }
        //parseProjectDetailsData();
    }

    /* Creating XLS File from CSV File downloaded*/
    private void parseData() {

        try {

            File csvfile = new File(dataFolder + "\\cmt_case_data_V2.csv");

            HSSFWorkbook workBook = new HSSFWorkbook();
            String xlsFileAddress = dataFolder + "\\cmt_case_data_V2.xls";
            HSSFSheet sheet = workBook.createSheet("Data");
            CreationHelper helper = workBook.getCreationHelper();

            int r = 0;

            CsvParserSettings settings = new CsvParserSettings();
            settings.setMaxCharsPerColumn(100000);
            settings.getFormat().setLineSeparator("\n");

            CsvParser parser = new CsvParser(settings);
            parser.beginParsing(csvfile);

            String[] row;

            while ((row = parser.parseNext()) != null) {

                Row frow = sheet.createRow((short) r++);
                for (int i = 0; i <row.length ; i++) {
                    frow.createCell(i).setCellValue(helper.createRichTextString(row[i]));
                }
            }

            parser.stopParsing();

            int lastRow = sheet.getLastRowNum();
            for (int i = 0; i < 7; i++) {
                sheet.removeRow(sheet.getRow(lastRow - i));
            }

            FileOutputStream fileOutputStream = new FileOutputStream(xlsFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();

        }catch (Exception e) {
            logger.log(Level.WARNING, "Parse Case Data Failed! Refer to Exception", e);
        }

        /*try {
            File csvfile = new File(System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data_V2.csv");

            HSSFWorkbook workBook = new HSSFWorkbook();

            String xlsFileAddress = System.getProperty("user.home") + "\\Documents\\CMT\\cmt_case_data_V2.xls";
            HSSFSheet sheet = workBook.createSheet("Data");

            BufferedReader br = new BufferedReader(new FileReader(csvfile));
            String line;

            int RowNum = 0;

            while ((line = br.readLine()) != null) {
                line = line.replaceAll("^\"|\"$", "");
                line = line.replaceAll(".0000000000", "");

                String[] fields = parseCsvLine(line);

                HSSFRow currentRow = sheet.createRow(RowNum);
                for (int i = 0; i < fields.length; i++) {
                    currentRow.createCell(i).setCellValue(fields[i]);
                    if (currentRow.getCell(i).getStringCellValue().isEmpty() || currentRow.getCell(i).getStringCellValue().equals("FALSE")) {
                        currentRow.getCell(i).setCellValue("NotSet");
                    }
                }
                RowNum++;
            }

            int lastRow = sheet.getLastRowNum();

            for (int i = 0; i < 7; i++) {
                sheet.removeRow(sheet.getRow(lastRow - i));
            }

            FileOutputStream fileOutputStream = new FileOutputStream(xlsFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }*/
    }
    private void rectifyAccountNames(){

        HSSFCell account;
        HSSFCell cellVal;
        HSSFCell age;

        try (HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(dataFolder + "\\cmt_case_data_V2.xls")))) {
            HSSFSheet filtersheet = workbook.getSheetAt(0);

            int cellnum = filtersheet.getRow(0).getLastCellNum();
            int lastRow = filtersheet.getLastRowNum();

            int row = 0;

            for (int i = 0; i < cellnum; i++) {

                String filterColName = filtersheet.getRow(0).getCell(i).toString();

                switch (filterColName) {
                    case ("Account Name"):
                        caseNumCellRefData = i;
                        break;
                    case ("Age (Days)"):
                        mycaseAgeRefCellData = i;
                        break;
                }
            }

            for (int i = 1; i < lastRow + 1; i++) {

                account = filtersheet.getRow(i).getCell(caseNumCellRefData);
                String caseStatus = account.getStringCellValue();
                caseStatus = caseStatus.replace(",", "");
                account.setCellValue(caseStatus);

                age = filtersheet.getRow(i).getCell(mycaseAgeRefCellData);
                String ageVal = age.getStringCellValue();
                ageVal = ageVal.replace(".0000000000", "");
                age.setCellValue(ageVal);

                for (int j = 0; j < cellnum; j ++){

                    cellVal = filtersheet.getRow(i).getCell(j);
                    String cellValue = cellVal.getStringCellValue();

                    if (cellValue.equals("")){
                        cellValue = "NotSet";
                        cellVal.setCellValue(cellValue);
                    }
                }
            }

            FileOutputStream output_file =new FileOutputStream(new File(dataFolder + "\\cmt_case_data_V3.xls"));
            workbook.write(output_file);
            output_file.close();

        }catch (Exception e){
            logger.log(Level.WARNING, "Rectify account names failed, please refer to exception", e);
        }
    }
}
