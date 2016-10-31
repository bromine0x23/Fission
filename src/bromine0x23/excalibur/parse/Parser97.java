package bromine0x23.excalibur.parse;

import bromine0x23.excalibur.Key;
import lombok.*;
import org.apache.poi.hssf.eventmodel.ERFListener;
import org.apache.poi.hssf.eventmodel.EventRecordFactory;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.record.common.UnicodeString;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.hssf.record.BOFRecord.TYPE_WORKSHEET;

@RequiredArgsConstructor()
public class Parser97 implements IParser {

	@NoArgsConstructor(access = AccessLevel.PRIVATE)
	private class Listener implements ERFListener {

		private List<UnicodeString> sst = new ArrayList<>();

		private List<String> sheetNames = new ArrayList<>();

		private int sheetId = -1;

		@Override
		public boolean processRecord(Record record) {
			switch (record.getSid()) {
				case BoundSheetRecord.sid:
					return handle((BoundSheetRecord)record);
				case BOFRecord.sid:
					return handle((BOFRecord)record);
				case SSTRecord.sid:
					return handle((SSTRecord)record);
				case LabelSSTRecord.sid:
					return handle((LabelSSTRecord)record);
				default:
					break;
			}
			return true;
		}

		private boolean handle(BoundSheetRecord record) {
			sheetNames.add(record.getSheetname());
			return true;
		}

		private boolean handle(BOFRecord record) {
			if (record.getType() == TYPE_WORKSHEET) {
				++sheetId;
			}
			return true;
		}

		private boolean handle(SSTRecord record) {
			for (int i = 0; i < record.getNumUniqueStrings(); ++i) {
				sst.add(record.getString(i));
			}
			return true;
		}

		private boolean handle(LabelSSTRecord record) {
			if (record.getRow() < 1) {
				keys.add(new Key(record.getColumn(), sst.get(record.getSSTIndex()).getString(), sheetId, sheetNames.get(sheetId)));
			}
			return true;
		}
	}

	@SuppressWarnings("MismatchedQueryAndUpdateOfCollection")
	@NonNull
	private List<Key> keys;

	private Listener listener = new Listener();

	@SuppressWarnings("MismatchedReadAndWriteOfArray")
	private static final short[] sids = new short[]{
		BoundSheetRecord.sid,
		BOFRecord.sid,
		SSTRecord.sid,
		LabelSSTRecord.sid
	};

	@Override
	public void parse(File file) throws Exception {
		keys.clear();
		val factory = new EventRecordFactory(listener, sids);
		try (val fileStream = new FileInputStream(file)) {
			val system = new NPOIFSFileSystem(fileStream);
			try (val stream = system.createDocumentInputStream("Workbook")) {
				factory.processRecords(stream);
			}
		}
	}

	public static void main(String args[]) throws Exception {
		final val path = "D:\\workspace\\Java\\Excalibur\\sample\\sample.xls";

		List<Key> keys = new ArrayList<>();

		val parser = new Parser97(keys);

		parser.parse(new File(path));

		System.out.printf("%d\n", keys.size());
		for (val key : keys) {
			System.out.println(key);
		}
	}
}
