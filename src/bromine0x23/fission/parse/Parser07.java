package bromine0x23.fission.parse;

import bromine0x23.excalibur.Key;
import bromine0x23.excalibur.Utility;
import lombok.*;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@RequiredArgsConstructor()
public class Parser07 implements IParser {

	private static class ParseTerminateException extends RuntimeException {
	}

	@NoArgsConstructor(access = AccessLevel.PRIVATE)
	private class Handler implements XSSFSheetXMLHandler.SheetContentsHandler {

		private int row;

		@Getter
		@Setter(AccessLevel.PRIVATE)
		private int sheetId = -1;

		@Getter
		@Setter(AccessLevel.PRIVATE)
		private String sheetName = "";

		@Override
		public void startRow(int rowNum) {
			row = rowNum;
		}

		@Override
		public void endRow(int rowNum) throws ParseTerminateException {
			throw new ParseTerminateException();
		}

		@Override
		public void cell(String cellReference, String formattedValue, XSSFComment comment) {
			if (row < 1) {
				int columnId = Utility.cellReferenceToInt(cellReference);
				keys.add(new Key(columnId, formattedValue, sheetId, sheetName));
			}
		}

		@Override
		public void headerFooter(String text, boolean isHeader, String tagName) {
		}
	}

	@SuppressWarnings("MismatchedQueryAndUpdateOfCollection")
	@NonNull
	private List<Key> keys;

	private Handler handler = new Handler();

	@Override
	public void parse(File file) throws Exception {
		keys.clear();

		val pkg = OPCPackage.open(file);
		val xssfReader = new XSSFReader(pkg);
		ContentHandler contentHandler = new XSSFSheetXMLHandler(xssfReader.getStylesTable(), new ReadOnlySharedStringsTable(pkg), handler, true);
		val xmlReader = XMLReaderFactory.createXMLReader();
		xmlReader.setContentHandler(contentHandler);
		val sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
		while (sheets.hasNext()) {
			try (val sheet = sheets.next()) {
				handler.setSheetId(handler.getSheetId() + 1);
				handler.setSheetName(sheets.getSheetName());
				val sheetSource = new InputSource(sheet);
				try {
					xmlReader.parse(sheetSource);
				} catch (ParseTerminateException e) {
					// ignore
				}
			}
		}
	}

	public static void main(String args[]) throws Exception {
		final val path = "D:\\workspace\\Java\\Excalibur\\sample\\9月固网出账明细（陈倩）.xlsx";

		List<Key> keys = new ArrayList<>();

		val parser = new Parser07(keys);

		parser.parse(new File(path));

		System.out.printf("%d\n", keys.size());
		for (val key : keys) {
			System.out.println(key);
		}
	}
}
