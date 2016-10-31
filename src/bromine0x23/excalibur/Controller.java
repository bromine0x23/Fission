package bromine0x23.excalibur;

import bromine0x23.excalibur.export.Exporter07;
import bromine0x23.excalibur.export.IExporter;
import bromine0x23.excalibur.parse.ParserProxy;
import bromine0x23.excalibur.pick.Picker07;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

import java.io.File;
import java.nio.file.Paths;

public class Controller {

	@FXML public TextField fieldInput;
	@FXML public TextField fieldOutput;
	@FXML public ComboBox<Key> boxKeys;
	@FXML public Button buttonInput;
	@FXML public Button buttonOutput;
	@FXML public Button buttonExport;
	@FXML public TextArea console;

	private ObservableList<Key> keys = FXCollections.observableArrayList();

	private File inputFile = null;
	private File outputDirectory = null;

	@SuppressWarnings("unused")
	@FXML public void initialize() {
		boxKeys.setItems(keys);
		writeLog("#");
		writeLog("# 标题行须位于第一行");
		writeLog("#");
	}

	@FXML public void buttonInputClick(@SuppressWarnings("UnusedParameters") MouseEvent mouseEvent) throws Exception {
		FileChooser chooser = new FileChooser();
		chooser.setTitle("选择Excel文件");
		chooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel 工作簿", "*.xls;*.xlsx"));
		chooser.setInitialDirectory(new File(System.getProperty("user.dir")));
		inputFile = chooser.showOpenDialog(null);
		if (inputFile != null) {
			fieldInput.setText(inputFile.getPath());

			ParserProxy.parse(keys, inputFile);
			if (keys.size() > 0) {
				boxKeys.setValue(keys.get(0));
			}
			boxKeys.setDisable(false);

			writeLog("识别了 %s 个候选主键列;", keys.size());

			if (outputDirectory == null) {
				outputDirectory = Paths.get(inputFile.getParent(), "拆分结果").toFile();
				fieldOutput.setText(outputDirectory.getPath());
			}
		}
	}

	@FXML public void buttonOutputClick(@SuppressWarnings("UnusedParameters") MouseEvent mouseEvent) {
		DirectoryChooser chooser = new DirectoryChooser();
		chooser.setTitle("选择导出文件夹");
		chooser.setInitialDirectory(inputFile.getParentFile());
		outputDirectory = chooser.showDialog(null);
		if (outputDirectory != null) {
			fieldOutput.setText(outputDirectory.getPath());
		}
	}

	@FXML public void buttonExportClick(@SuppressWarnings("UnusedParameters") MouseEvent mouseEvent) throws Exception {
		if (inputFile == null) {
			new Alert(Alert.AlertType.ERROR, "未选择输入文件", ButtonType.CLOSE).show();
			return;
		}
		if (outputDirectory == null) {
			new Alert(Alert.AlertType.ERROR, "未指定输出文件夹", ButtonType.CLOSE).show();
			return;
		}
		Key key = boxKeys.getValue();
		if (key == null) {
			new Alert(Alert.AlertType.ERROR, "未指定主键列", ButtonType.CLOSE).show();
			return;
		}

		IExporter exporter = new Exporter07();

		Picker07 picker07 = new Picker07(key, exporter);

		writeLog("读取Excel表……");

		picker07.parse(inputFile);

		writeLog("导出拆分结果……");

		int count = exporter.export(outputDirectory, inputFile.getName());

		writeLog("导出了 %d 个子表到目录 %s ;", count, outputDirectory.getPath());

		writeLog("共 %d 列（不计标题行）;", exporter.getColumnCount());
	}

	private void writeLog(String format, Object ... args) {
		console.appendText(String.format(format, args));
		console.appendText("\n");
	}
}
