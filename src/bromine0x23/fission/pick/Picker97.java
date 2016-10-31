package bromine0x23.fission.pick;

import bromine0x23.excalibur.Key;
import bromine0x23.excalibur.export.Exporter07;
import bromine0x23.excalibur.export.IExporter;
import bromine0x23.excalibur.parse.Parser97;
import bromine0x23.excalibur.sheet.Cell;
import lombok.*;
import org.apache.poi.hssf.eventmodel.ERFListener;
import org.apache.poi.hssf.eventmodel.EventRecordFactory;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.record.common.UnicodeString;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.apache.poi.hssf.record.BOFRecord.TYPE_WORKSHEET;

@RequiredArgsConstructor()
public class Picker97 implements IPicker {

	@NoArgsConstructor(access = AccessLevel.PRIVATE)
	private class Listener implements ERFListener {

		private List<UnicodeString> sst = new ArrayList<>();

		private Map<Integer, Cell> values = new HashMap<>();

		private int sheetId = -1;

		private int row = 0;

		private int column;

		@Override
		public boolean processRecord(Record record) {
			switch (record.getSid()) {
				case BOFRecord.sid:
					return handle((BOFRecord)record);
				case SSTRecord.sid:
					return handle((SSTRecord)record);
				case LabelSSTRecord.sid:
					return handle((LabelSSTRecord)record);
				case NumberRecord.sid:
					return handle((NumberRecord)record);
				case BoolErrRecord.sid:
					return handle((BoolErrRecord)record);
				case FormulaRecord.sid:
					return handle((FormulaRecord)record);
				default:
					break;
			}
			return true;
		}

		private boolean handle(BOFRecord record) {
			if (record.getType() != TYPE_WORKSHEET) {
				++sheetId;
			}
			return sheetId <= key.getSheetId();
		}

		private boolean handle(SSTRecord record) {
			for (int i = 0; i < record.getNumUniqueStrings(); ++i) {
				sst.add(record.getString(i));
			}
			return true;
		}

		private boolean handle(LabelSSTRecord record) {
			handleCell(record);
			String value = sst.get(record.getSSTIndex()).getString();
			values.put(column, Cell.createStringCell(value));
			return true;
		}

		private boolean handle(NumberRecord record) {
			handleCell(record);
			values.put(column, Cell.createNumericCell(record.getValue()));
			return true;
		}

		private boolean handle(BoolErrRecord record) {
			handleCell(record);
			Cell cell;
			if (record.isBoolean()) {
				cell = Cell.createBooleanCell(record.getBooleanValue());
			} else {
				assert record.isError();
				cell = Cell.createErrorCell(record.getErrorValue());
			}
			values.put(column, cell);
			return true;
		}

		@SuppressWarnings("deprecation")
		private boolean handle(FormulaRecord record) {
			handleCell(record);
			Cell cell;
			if (record.getCachedResultType() == CellType.BOOLEAN.getCode()) {
				cell = Cell.createBooleanCell(record.getCachedBooleanValue());
			} else if (record.getCachedResultType() == CellType.ERROR.getCode()) {
				cell = Cell.createErrorCell((byte)record.getCachedErrorValue());
			} else {
				cell = Cell.createNumericCell(record.getValue());
			}
			values.put(column, cell);
			return true;
		}

		private void handleCell(CellRecord record) {
			column = record.getColumn();
			if (row < record.getRow()) {
				if (row == 0) {
					exporter.feedHeader(values);
				} else {
					exporter.feed(values.get(key.getColumnId()).asString(), values);
				}
				values.clear();
				row = record.getRow();
			}
		}
	}

	@NonNull
	private Key key;

	@NonNull
	private IExporter exporter;

	private Listener listener = new Listener();

	@SuppressWarnings("MismatchedReadAndWriteOfArray")
	private static final short[] sids = new short[]{
		BoundSheetRecord.sid,
		BOFRecord.sid,
		SSTRecord.sid,
		LabelSSTRecord.sid,
		NumberRecord.sid,
		BoolErrRecord.sid,
		FormulaRecord.sid,
	};

	@Override
	public void parse(File file) throws Exception {
		val factory = new EventRecordFactory(listener, sids);
		try (val fileStream = new FileInputStream(file)) {
			val system = new NPOIFSFileSystem(fileStream);
			try (val stream =  system.createDocumentInputStream("Workbook")) {
				factory.processRecords(stream);
			}
		}
	}

	public static void main(String args[]) {
		final val inputPath = "D:\\workspace\\Java\\Excalibur\\sample\\big.xls";
		final val exportPath = "D:\\workspace\\Java\\Excalibur\\sample\\export";

		val inputFile = new File(inputPath);

		List<Key> keys = new ArrayList<>();

		val parser = new Parser97(keys);

		try {
			parser.parse(inputFile);
		} catch (Exception e) {
			e.printStackTrace();
		}

		val key = keys.get(0);

		System.out.printf("key = %s\n", key);

		IExporter exporter = new Exporter07();

		val picker = new Picker97(key, exporter);

		System.out.println("Parse");
		try {
			picker.parse(inputFile);
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("Export");
		try {
			exporter.export(new File(exportPath), inputFile.getName());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
