package bromine0x23.fission;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NonNull;

@AllArgsConstructor()
public class Key {

	@Getter
	private int columnId;

	@NonNull
	@Getter
	private String columnName;

	@Getter
	private int sheetId;

	@NonNull
	@Getter
	private String sheetName;

	@Override
	public String toString() {
		return String.format("%s (%s #%s)", columnName, sheetName, sheetId + 1);
	}
}
