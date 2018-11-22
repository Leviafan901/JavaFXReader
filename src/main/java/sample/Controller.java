package sample;

import javafx.embed.swing.SwingFXUtils;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.commons.io.IOUtils;
import org.apache.poi.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.imageio.ImageIO;
import java.io.*;
import java.nio.file.Files;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class Controller {

    @FXML
    private TextField firstName;

    @FXML
    private Button generateDocButton;

    private Map<String, String> values = new HashMap<>();

    public void initialize() {
        generateDocButton.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                String initPath = "D:\\github\\readerIK\\src\\main\\resources\\docs\\Test.docx";
                try (FileInputStream file = new FileInputStream(new File(initPath));
                     XWPFDocument docx = new XWPFDocument(OPCPackage.open(file))) {
                    replaceParagraph(docx, "reason", firstName.getText());
                    final FileOutputStream out = new FileOutputStream(String.format("D:/%s.docx", firstName.getText()));
                    docx.write(out);
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
