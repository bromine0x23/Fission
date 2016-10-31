package bromine0x23.fission.pick;

import bromine0x23.excalibur.Key;
import bromine0x23.excalibur.Utility;
import bromine0x23.excalibur.export.Exporter07;
import bromine0x23.excalibur.export.IExporter;
import bromine0x23.excalibur.parse.Parser07;
import bromine0x23.excalibur.sheet.Cell;
import lombok.NonNull;
import lombok.RequiredArgsConstructor;
import lombok.val;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.InputStream;
import java.util.*;

@RequiredArgsConstructor()
public class Picker07 implements IPicker {

	private enum DataType { Boolean, Error, Formula, InlineString, SSTIndex, Numeric }

	@RequiredArgsConstructor()
	private class Handler extends DefaultHandler {

		private Map<Integer, Cell> values = new HashMap<>();

		@NonNull
		private StylesTable stylesTable;

		@NonNull
		private SharedStringsTable sharedStringsTable;

		private boolean vOpened;

		private DataType nextDataType = DataType.Numeric;

		private boolean firstRow = true;
		private int column = -1;
		private int styleIndex;

		private StringBuffer valueBuffer = new StringBuffer();

		@Override
		public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
			 if ("c".equals(qName)) {
				this.column = Utility.cellReferenceToInt(attributes.getValue("r"));

				this.nextDataType = DataType.Numeric;
				 val cellType = attributes.getValue("t");
				 val styleIndex = attributes.getValue("s");
				if (styleIndex != null) {
					this.styleIndex = Integer.parseInt(styleIndex);
				} else {
					this.styleIndex = -1;
				}

				if ("b".equals(cellType)) {
					nextDataType = DataType.Boolean;
				} else if ("e".equals(cellType)) {
					nextDataType = DataType.Error;
				} else if ("inlineStr".equals(cellType)) {
					nextDataType = DataType.InlineString;
				} else if ("s".equals(cellType)) {
					nextDataType = DataType.SSTIndex;
				} else if ("str".equals(cellType)) {
					nextDataType = DataType.Formula;
				}
			} else if ("inlineStr".equals(qName) || "v".equals(qName)) {
				 vOpened = true;
				 valueBuffer.setLength(0);
			 } else if ("row".equals(qName)) {
				values.clear();
			} else if ("worksheet".equals(qName)) {
				exporter.feedStyles(stylesTable);
			}
		}

		@Override
		public void endElement(String uri, String localName, String qName) throws SAXException {
			if ("v".equals(qName)) {
				Cell cell = null;
				switch (nextDataType) {
					case Boolean: {
						cell = Cell.createBooleanCell(valueBuffer.charAt(0) == '0');
						break;
					}
					case Error: {
						cell = Cell.createStringCell(valueBuffer.toString());
						break;
					}
					case Formula: {
						String value = valueBuffer.toString();
						try {
							cell = Cell.createNumericCell(Double.parseDouble(value));
						} catch(NumberFormatException e) {
							cell = Cell.createStringCell(value);
						}
						break;
					}
					case InlineString: {
						XSSFRichTextString textString = new XSSFRichTextString(valueBuffer.toString());
						cell = Cell.createStringCell(textString.toString());
						break;
					}
					case SSTIndex: {
						try {
							int index = Integer.parseInt(valueBuffer.toString());
							XSSFRichTextString text = new XSSFRichTextString(sharedStringsTable.getEntryAt(index));
							cell = Cell.createStringCell(text.toString());
						} catch (NumberFormatException ex) {
							// ignored
						}
						break;
					}
					case Numeric: {
						double value = Double.parseDouble(valueBuffer.toString());
						cell = Cell.createNumericCell(value);
						break;
					}
				}
				if (styleIndex >= 0) {
					assert cell != null;
					cell.setStyleIndex(styleIndex);
				}
				values.put(column, cell);
			} else if ("row".equals(qName)) {
				if (firstRow) {
					firstRow = false;
					exporter.feedHeader(values);
				} else {
					exporter.feed(values.get(key.getColumnId()).asString(), values);
				}
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) throws SAXException {
			if (vOpened) {
				valueBuffer.append(ch, start, length);
			}
		}
	}

	@NonNull
	private Key key;

	@NonNull
	private IExporter exporter;

	@Override
	public void parse(File file) throws Exception {
		val pkg = OPCPackage.open(file);
		val xssfReader = new XSSFReader(pkg);
		val handler = new Handler(xssfReader.getStylesTable(), xssfReader.getSharedStringsTable());
		val xmlReader = XMLReaderFactory.createXMLReader();
		xmlReader.setContentHandler(handler);
		int sheetId = 0;
		val sheets = xssfReader.getSheetsData();
		while (sheets.hasNext()) {
			try (val sheet = sheets.next()) {
				if (sheetId == key.getSheetId()) {
					InputSource sheetSource = new InputSource(sheet);
					xmlReader.parse(sheetSource);
					break;
				}
				++sheetId;
			}
		}
	}

	public static void main(String args[]) {
		final val inputPath = "D:\\workspace\\Java\\Excalibur\\sample\\9月固网出账明细（陈倩）.xlsx";
		final val exportPath = "D:\\workspace\\Java\\Excalibur\\sample\\export";

		val inputFile = new File(inputPath);

		List<Key> keys = new ArrayList<>();

		val parser = new Parser07(keys);

		try {
			parser.parse(inputFile);
		} catch (Exception e) {
			e.printStackTrace();
		}

		val key = new Key(13, "团队", 1, "");

		// Key key = keys.get(0);

		System.out.printf("key = %s\n", key);

		IExporter exporter = new Exporter07();

		val picker = new Picker07(key, exporter);

		System.out.println("Parse");
		try {
			picker.parse(inputFile);
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("Export");
		try {
			int count = exporter.export(new File(exportPath), inputFile.getName());
			System.out.printf("Total Sheets: %d\n", count);
			System.out.printf("Total Columns: %d\n", exporter.getColumnCount());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
