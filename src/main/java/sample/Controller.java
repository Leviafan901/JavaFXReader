package sample;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.*;

public class Controller {

    @FXML
    private TextField name;

    @FXML
    private TextField birthDate;

    @FXML
    private TextField birthPlace;

    @FXML
    private TextField registrationAddress;

    @FXML
    private Label personalIDLabel = new Label("Личный номер \n (для имущественных запросов)");

    @FXML
    private TextField personalID;

    @FXML
    private TextField detective;

    @FXML
    private TextField phoneNumber;

    @FXML
    private Label firstLastNameLabel;

    @FXML
    private TextField orderNumber;

    @FXML
    private TextField stateNumber;

    @FXML
    private ChoiceBox<String> requestTypeChoiceBox;

    @FXML
    private ChoiceBox<String> requestBaseChoiceBox;

    @FXML
    private Button generateDocButton;

    private final Map<String, String> values = new HashMap<>();

    public void initialize() {

        initChoiceBox(requestTypeChoiceBox, "Запросы в диспансеры", "Имущественные запросы");
        initChoiceBox(requestBaseChoiceBox, "Материал проверки КР", "Уголовное дело");

        generateDocButton.setOnAction(event -> {
            defineRequestBase(requestBaseChoiceBox);
            generateFile(requestType(requestTypeChoiceBox), initValuesMap());
        });
    }

    private void generateFile(String filePath, Map<String, String> keyWords) {
        try (FileInputStream file = new FileInputStream(filePath);
             XWPFDocument docx = new XWPFDocument(OPCPackage.open(file))) {
            replaceParagraph(docx, keyWords);
            replaceTextInFooter(docx, keyWords);
            FileChooser fileChooser = new FileChooser();
            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Word files (*.docx)", "*.docx");
            fileChooser.getExtensionFilters().add(extFilter);
            Stage wind = new Stage();
            File newFile = fileChooser.showSaveDialog(wind);
            if (newFile != null) {
                final FileOutputStream outF = new FileOutputStream(newFile);
                docx.write(outF);
            }
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private void replaceParagraph(XWPFDocument doc, Map<String, String> keyWords) {
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    for (String key : keyWords.keySet()) {
                        String text = r.getText(0);
                        if (text != null && text.contains(key)) {
                            text = text
                                    .replaceAll(key, keyWords.get(key));
                            r.setText(text, 0);
                        }
                    }
                }
            }
        }
    }

    private void replaceTextInFooter(XWPFDocument doc, Map<String, String> keyWords) {
        for (XWPFFooter footer : doc.getFooterList()) {
            for (XWPFParagraph paragraph : footer.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    if (run != null) {
                        for (String key : keyWords.keySet()) {
                            String text = run.getText(0);
                            if (text != null && text.contains(key)) {
                                text = text
                                        .replaceAll(key, keyWords.get(key));
                                run.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
    }

    private Map<String, String> initValuesMap() {
        values.put("name", textFieldValue(name));
        values.put("birthDate", textFieldValue(birthDate));
        values.put("birthPlace", textFieldValue(birthPlace));
        values.put("registration", textFieldValue(registrationAddress));
        values.put("orderNumber", textFieldValue(orderNumber));
        values.put("stateNumber", textFieldValue(stateNumber));
        values.put("personalID", textFieldValue(personalID));
        values.put("detective", textFieldValue(detective));
        values.put("phoneNumber", textFieldValue(phoneNumber));
        return values;
    }

    private String textFieldValue(TextField field) {
        String value = field.getText();
        return (value == (null))
                ? ""
                : value;
    }

    private void initChoiceBox(ChoiceBox<String> choiceBox, String... args) {
        List<String> typeList = new ArrayList<>(Arrays.asList(args));
        ObservableList<String> types = FXCollections.observableList(typeList);
        choiceBox.setItems(types);
    }

    private String requestType(ChoiceBox<String> typesChoiceBox) {
        return switch (typesChoiceBox.getValue()) {
            case "Имущественные запросы" -> "docs/PropertyRequests.docx";
            default -> "docs/PublicHealth.docx";
        };
    }

    private void defineRequestBase(ChoiceBox<String> typesChoiceBox) {
        switch (typesChoiceBox.getValue()) {
            case "Материал проверки КР" -> {
                values.put("criminal", "");
                values.put("requestBase", "рассмотрением материалов проверки КР №");
            }
            case "Уголовное дело" -> {
                values.put("criminal", "возбужденного по ");
                values.put("requestBase", "расследованием уголовного дела №");
            }
            default -> {
                values.put("criminal", "");
                values.put("reason", "рассмотрением материалов проверки КР №");
            }
        }
    }
}
