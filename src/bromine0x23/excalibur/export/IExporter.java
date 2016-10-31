package bromine0x23.excalibur.export;

import bromine0x23.excalibur.sheet.Cell;
import org.apache.poi.xssf.model.StylesTable;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.util.Map;

public interface IExporter extends Closeable {

	void feedStyles(StylesTable stylesTable);

	void feedHeader(Map<Integer, Cell> values);

	void feed(String key, Map<Integer, Cell> values);

	int export(File outputDirectory, String inputFilename) throws IOException;

	int getColumnCount();
}
