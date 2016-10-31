package bromine0x23.excalibur.sheet;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.xssf.streaming.SXSSFCell;

@NoArgsConstructor(access = AccessLevel.PRIVATE)
public class Cell {
	private enum Type {
		Numeric, String, Boolean, Error
	}

	private Type type;

	@Getter
	private double numericValue;

	@Getter
	private String stringValue;

	@Getter
	private boolean booleanValue;

	@Getter
	private byte errorValue;

	@Getter
	@Setter
	private int styleIndex = -1;

	public static Cell createNumericCell(double value) {
		Cell cell = new Cell();
		cell.setNumericValue(value);
		return cell;
	}

	public static Cell createStringCell(String value) {
		Cell cell = new Cell();
		cell.setStringValue(value);
		return cell;
	}

	public static Cell createBooleanCell(boolean value) {
		Cell cell = new Cell();
		cell.setBooleanValue(value);
		return cell;
	}

	public static Cell createErrorCell(byte value) {
		Cell cell = new Cell();
		cell.setErrorValue(value);
		return cell;
	}

	private void setNumericValue(double value) {
		numericValue = value;
		type = Type.Numeric;
	}

	private void setStringValue(String value) {
		stringValue = value;
		type = Type.String;
	}

	private void setBooleanValue(boolean value) {
		booleanValue = value;
		type = Type.Boolean;
	}

	private void setErrorValue(byte value) {
		errorValue = value;
		type = Type.Error;
	}

	public String asString() {
		switch (type) {
			case Numeric:
				return Double.toString(numericValue);
			case String:
				return stringValue;
			case Boolean:
				return Boolean.toString(booleanValue);
			case Error:
				return Byte.toString(errorValue);
		}
		assert false;
		return "";
	}

	public void copyTo(SXSSFCell cell) {
		switch (type) {
			case Numeric:
				cell.setCellValue(numericValue);
				break;
			case String:
				cell.setCellValue(stringValue);
				break;
			case Boolean:
				cell.setCellValue(booleanValue);
				break;
			case Error:
				cell.setCellErrorValue(errorValue);
				break;
		}

	}
}
