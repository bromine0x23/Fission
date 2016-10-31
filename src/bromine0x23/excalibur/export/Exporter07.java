package bromine0x23.excalibur.export;

import bromine0x23.excalibur.Utility;
import bromine0x23.excalibur.sheet.Cell;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.val;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.util.*;

@NoArgsConstructor()
public class Exporter07 implements IExporter {

	private static final int ROW_ACCESS_WINDOW_SIZE = 50;

	private static class Bind implements Closeable {
		SXSSFWorkbook workbook;
		SXSSFSheet sheet;
		Map<Integer, Short> styleIndexMap = new HashMap<>();
		int rowCount = 0;

		private Bind(StylesTable stylesTable, Map<Integer, Cell> header) {
			this.workbook = new SXSSFWorkbook(null, ROW_ACCESS_WINDOW_SIZE, false, true);
			this.sheet = workbook.createSheet();
			if (stylesTable != null) {
				for (int i = 0; i < stylesTable.getNumCellStyles(); ++i) {
					val style = workbook.createCellStyle();
					style.cloneStyleFrom(stylesTable.getStyleAt(i));
					styleIndexMap.put(i, style.getIndex());
				}
			}
			appendRow(header);
		}

		private CellStyle getStyle(int styleIndex) {
			return workbook.getCellStyleAt((int)styleIndexMap.get(styleIndex));
		}

		private void appendRow(Map<Integer, Cell> values) {
			val row = sheet.createRow(rowCount++);
			for (val entry : values.entrySet()) {
				val cell = entry.getValue();
				val sxssfCell = row.createCell(entry.getKey());
				if (cell.getStyleIndex() >= 0) {
					sxssfCell.setCellStyle(getStyle(cell.getStyleIndex()));
				}
				cell.copyTo(sxssfCell);
			}
		}

		@Override
		public void close() throws IOException {
			workbook.close();
			workbook.dispose();
		}
	}

	private Map<String, Bind> binds = new HashMap<>();

	private Map<Integer, Cell> header;

	private StylesTable stylesTable = null;

	@Getter
	private int columnCount = 0;

	@Override
	public void feedStyles(StylesTable stylesTable) {
		this.stylesTable = stylesTable;
	}

	@Override
	public void feedHeader(Map<Integer, Cell> header) {
		this.header = new HashMap<>(header);
	}

	@Override
	public void feed(String key, Map<Integer, Cell> values) {
		getBind(key).appendRow(values);
		++columnCount;
	}

	@SuppressWarnings("ResultOfMethodCallIgnored")
	@Override
	public int export(File outputDirectory, String inputFilename) throws IOException {
		outputDirectory.mkdirs();
		for (val entry : binds.entrySet()) {
			val key = entry.getKey().replaceAll("[/:*?\"<>|\\\\]", "_");
			val exportFilename = String.format("%s_%s.xlsx", Utility.getBaseName(inputFilename), key);
			val file = new File(outputDirectory, exportFilename);
			try (val stream = new FileOutputStream(file)) {
				entry.getValue().workbook.write(stream);
			}
		}
		return binds.size();
	}

	@Override
	public void close() throws IOException {
		for (val bind : binds.values()) {
			bind.close();
		}
		binds.clear();
	}

	private Bind getBind(String columnValue) {
		if (binds.containsKey(columnValue)) {
			return binds.get(columnValue);
		} else {
			val bind = new Bind(stylesTable, header);
			binds.put(columnValue, bind);
			return bind;
		}
	}

	public static void main(String[] args) throws IOException {
		Random random = new Random();

		SXSSFWorkbook workbook = new SXSSFWorkbook(1024);
		workbook.createCellStyle();
		SXSSFSheet sheet = workbook.createSheet();

		final int maxRow = 0x80000;
		final int maxCol = 0x20;

		{
			SXSSFRow headerRow = sheet.createRow(0);
			for (int c = 0; c < maxCol; ++c) {
				SXSSFCell cell = headerRow.createCell(c);
				cell.setCellValue(String.format("Column#%d", c));
			}
		}

		final String ALPHABETA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

		for (int r = 1; r < maxRow; ++r) {
			SXSSFRow row = sheet.createRow(r);
			int c = 0;
			{
				SXSSFCell cell = row.createCell(c++);
				cell.setCellValue(Character.toString(ALPHABETA.charAt(random.nextInt(ALPHABETA.length()))));
			}
			for (; c < maxCol; ++c) {
				SXSSFCell cell = row.createCell(c);
				cell.setCellValue(random.nextDouble() * random.nextDouble());
			}
		}

		try (OutputStream stream = new FileOutputStream("D:\\workspace\\Java\\Excalibur\\sample\\big.xlsx")) {
			workbook.write(stream);
		}
	}
}
