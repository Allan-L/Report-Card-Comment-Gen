package sample;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.List;

public class ReportCard {
    private FileOutputStream out;
    private XWPFDocument document;
    private XWPFParagraph paragraph;
    private XWPFRun run;
    private File file;
    private boolean exist;
    private boolean strongSkill;
    private boolean pass;
    private String sheetPath;
    private int skillColNum;


    public ReportCard() {
        sheetPath = "src/sample/ReportCards.xlsx";
    }

    public void setFile(File file) {
        this.file = file;

    }

    public void hasStrongSkill(String status){

        if(status.equals("Good")){
            strongSkill = true;
        }else{
            strongSkill = false;
        }



    }
    public void setStatus(String status){
        if(status.equals("Pass")){
            pass = true;
        }else{
            pass = false;
        }

    }


    public File getFile() {
        return file;
    }


    /**
     * This will get the opening comment for the report card
     * @return the opening comment for the report card
     */
    public String getOpeningComment() throws FileNotFoundException {
        int sheetNum = 0;
        String s = "";


        s = getSheet(sheetNum, 0);





        return s;

    }

    public int setComment(String comment){
        int colNum = -1;

        switch (comment){
            case "Front Float":
               colNum = 0;
                break;
            case "Back Float":
                colNum = 1;
                break;
            case "Front Glide":
                colNum = 2;
                break;
            case "Back Glide":
                colNum = 3;
                break;
            case "Side Glide":
                colNum = 4;
                break;
            case "Front Crawl":
                colNum = 5;
                break;
            case "Back Crawl":
                colNum = 6;
                break;
            case "Breast Stroke":
                colNum = 7;
                break;
            case "Whip Kick":
                colNum = 8;
                break;
            case "Flutter Kick":
                colNum = 9;
                break;
        }

        return colNum;

    }

    public int setSkillsComment(String comment){
        int colNum = -1;

        switch (comment){
            case "Threading Water":
                colNum = 0;
                break;
            case "Submerge Head":
                colNum = 1;
                break;
            case "Side Breathing":
                colNum = 2;
                break;
            case "Endurance":
                colNum = 3;
                break;
            case "Kneeling Dive":
                colNum = 4;
                break;
            case "Standing Dive":
                colNum = 5;
                break;
            case "Stride Jump":
                colNum = 6;
                break;



        }

        return colNum;

    }

    //Drop box Choices

    public String getPositiveComment(int colNum) throws FileNotFoundException {
        int sheetNum = 1;
        String s = "";
        s = getSheet(sheetNum, colNum);

        return s;
    }

    //Drop box Choices

    public String getCritiqueComment(int colNum) throws FileNotFoundException {
        int sheetNum = 2;
        String s = "";
        s = getSheet(sheetNum, colNum);

        return s;
    }

    //Drop box Choices w/ boolean skills for good or bad comment

    public String getSkillsComment(boolean skills, int colNum) throws FileNotFoundException {
        int sheetNum;
        String s = "";
        if (skills) {
            sheetNum = 3;
            s = getSheet(sheetNum, colNum);
        } else {
            sheetNum = 4;
            s = getSheet(sheetNum, colNum);
        }
        return s;

    }

    public String getFinalComment(Boolean pass) throws FileNotFoundException {
        int sheetNum = 5;
        String s = "";

        if (pass) {
            s = getSheet(sheetNum, 0);
        } else {
            s = getSheet(sheetNum, 1);
        }

        return s;

    }


    /**
     * This method will get the workbook and sheet
     *
     * @param sheet is which sheet in the workbook to access
     * @return a comment based on input.
     */
    public String getSheet(int sheetNum, int colNum) throws FileNotFoundException {
        String comment = "";


        try (Workbook wb = WorkbookFactory.create(new FileInputStream(sheetPath))) {

            Sheet sheet = wb.getSheetAt(sheetNum);

            int rowStart = sheet.getFirstRowNum();
            int rowEnd = sheet.getLastRowNum();
            for (int i = rowStart; i < rowEnd; i++) {
                //Get row of the sheet randomize row to get random comment plus one to ignore header
                int random = (int) (Math.random() * rowEnd);
                Row row = sheet.getRow(random + 1);




                if (row != null) {

                    //Get column of sheet
                    Cell cell = row.getCell(colNum);
                    if (cell != null) {
                        comment = cell.getStringCellValue();

                    }
                }
                wb.close();

            }
        } catch (IOException e) {
            e.printStackTrace();
        }


        return comment;
    }

    public void generateReportCard(String name,int posColNum, int negColNum, int skillColNum) throws IOException, InvalidFormatException {
        boolean fileCheck = !file.exists() || file.length() == 0;


        if (fileCheck) {
            document = new XWPFDocument();  //Get document file
            out = new FileOutputStream(file);  // Write to it
            paragraph = document.createParagraph();  //Make a paragraph
            run = paragraph.createRun(); //Add to paragraph
            run.setText(getOpeningComment() + " " + name + ". " + getPositiveComment(posColNum) + ". " + getCritiqueComment(negColNum) +". " +getSkillsComment(strongSkill,skillColNum) +". " +getFinalComment(pass) + " " + name + "."); //Set text of paragraph
            run.addBreak(); //Add a new line
            run.addBreak();
            document.write(out);
            document.close();
            out.close();

        } else if(file.exists()) {
            document = new XWPFDocument(OPCPackage.open(file)); //Get document file
            List<XWPFParagraph> paragraphs = document.getParagraphs(); //Get the paragraphs
            paragraph =  paragraphs.get(paragraphs.size() - 1);
            run = paragraph.createRun(); //Append new paragraph
            run.setText(getOpeningComment() + " " + name + ". " + getPositiveComment(posColNum) + ". " + getCritiqueComment(negColNum) +". " +getSkillsComment(strongSkill,skillColNum) +". " +getFinalComment(pass) + " " + name + ".");
            run.addBreak();
            run.addBreak();
            file.delete();


             //Read and Overwrite by deleting to prevent errors
            try ( FileOutputStream out = new FileOutputStream(file)) {

                document.write(out);

            } catch (IOException e) {

                e.printStackTrace();
            }
        }
        document.close();
        out.close();
}
    }

