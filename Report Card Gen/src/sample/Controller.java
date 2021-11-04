package sample;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.stage.FileChooser;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.TextField;
import javafx.scene.control.Label;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class Controller {




    @FXML
    private Button fileButton;
    @FXML
    private Button generateButton;
    @FXML
    private TextField nameField;
    @FXML
    private CheckBox passCheckBox;
    @FXML
    private CheckBox failCheckBox;
    @FXML
    private ChoiceBox<String> positiveChoice;
    @FXML
    private ChoiceBox<String> constructiveChoice;
    @FXML
    private ChoiceBox<String> skillsChoice;
    @FXML
    private CheckBox goodSkillCB;
    @FXML
    private CheckBox badSkillCB;
    @FXML
    void initialize() {


        positiveChoice.getItems().addAll("Front Float","Back Float","Front Glide","Back Glide","Side Glide","Front Crawl","Back Crawl","Breast Stroke","Whip Kick","Flutter Kick");

        constructiveChoice.getItems().addAll("Front Float","Back Float","Front Glide","Back Glide","Side Glide","Front Crawl","Back Crawl","Breast Stroke","Whip Kick","Flutter Kick");

        skillsChoice.getItems().setAll("Threading Water","Submerge Head","Side Breathing","Endurance", "Kneeling Dive", "Standing Dive", "Stride Jump");

        ReportCard reportCardGen = new ReportCard();


        fileButton.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                FileChooser chooser = new FileChooser();
                chooser.setTitle("Choose location To Save Report Card");
                File selectedFile = null;
                while(selectedFile == null){
                    selectedFile = chooser.showSaveDialog(null);
                }
                reportCardGen.setFile(new File(String.valueOf(selectedFile) + ".docx"));
            }
        });
        generateButton.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                if(passCheckBox.isSelected()){
                    reportCardGen.setStatus("Pass");
                    passCheckBox.setSelected(false);
                }
                if(failCheckBox.isSelected()){
                    reportCardGen.setStatus("Fail");
                    failCheckBox.setSelected(false);
                }
                if(goodSkillCB.isSelected()){
                    reportCardGen.hasStrongSkill("Good");
                    goodSkillCB.setSelected(false);
                }
                if (badSkillCB.isSelected()){
                    reportCardGen.hasStrongSkill("Bad");
                    badSkillCB.setSelected(false);
                }
                try {
            reportCardGen.generateReportCard(nameField.getText(),reportCardGen.setComment(positiveChoice.getValue()),reportCardGen.setComment(constructiveChoice.getValue()),reportCardGen.setSkillsComment(skillsChoice.getValue()));
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();

        }
                nameField.clear();
                positiveChoice.setValue(null);
                constructiveChoice.setValue(null);
                skillsChoice.setValue(null);
    }
});
    }
}
