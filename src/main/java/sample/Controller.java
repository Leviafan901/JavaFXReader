package sample;

import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Controller {

    @FXML
    private TextField firstName;

    @FXML
    private TextField birthDate;

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
    private ComboBox<?> requestTypeDropdown;

    @FXML
    private ComboBox<?> requestBaseDropdown;

    @FXML
    private Button generateDocButton;

    private Map<String, String> values = new HashMap<>();

    public void initialize() {
        generateDocButton.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                ClassLoader classLoader = getClass().getClassLoader();
                String initPath = classLoader.getResource("docs/PublicHealth.docx").getPath();
                try (FileInputStream file = new FileInputStream(new File(initPath));
                     XWPFDocument docx = new XWPFDocument(OPCPackage.open(file))) {
                    replaceParagraph(docx, "reason", firstName.getText());


                    FileChooser fileChooser = new FileChooser();
                    //Set extension filter for text files
                    FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Word files (*.docx)", "*.docx");
                    fileChooser.getExtensionFilters().add(extFilter);

                    Stage wind = new Stage();

                    //Show save file dialog
                    File newFile = fileChooser.showSaveDialog(wind);

                    if (newFile != null) {
                        final FileOutputStream outF = new FileOutputStream(newFile);
                        docx.write(outF);
                    }
                    //final FileOutputStream out = new FileOutputStream(String.format("D:/%s.docx", firstName.getText()));
                    //docx.write(out);
                } catch (IOException | InvalidFormatException e) {
                    e.printStackTrace();
                }
            }
        });
    }

    private void replaceParagraph(XWPFDocument doc, String key, String repl) {
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains(key)) {
                        text = text
                                .replaceAll(key, repl);
                        r.setText(text, 0);
                    }
                }
            }
        }
    }

    private Map<String, String> initValuesMap() {
        values.put("name", firstName.getText());
        return values;
    }
}
