package bromine0x23.fission;

public final class Utility {

	public static int cellReferenceToInt(String cellReference) {
		int column = -1;
		for (int i = 0; i < cellReference.length(); ++i) {
			int c = cellReference.charAt(i);
			if (Character.isDigit(c)) {
				break;
			} else {
				column = (column + 1) * 26 + c - 'A';
			}
		}
		return column;
	}

	public static String getBaseName(String filename) {
		int index = filename.lastIndexOf(".");
		return index == -1 ? filename : filename.substring(0, index);
	}
}
