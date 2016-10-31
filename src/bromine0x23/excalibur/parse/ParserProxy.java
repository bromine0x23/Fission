package bromine0x23.excalibur.parse;

import bromine0x23.excalibur.Key;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ParserProxy {

	public static void parse(List<Key> keys, File file) throws Exception {
		IParser parser;
		try (InputStream stream = new FileInputStream(file)) {
			parser = createParser(keys, stream);
		}
		parser.parse(file);
	}

	private static IParser createParser(List<Key> keys, InputStream stream) throws IOException {
		stream = new PushbackInputStream(stream, 4096);
		if (NPOIFSFileSystem.hasPOIFSHeader(stream)) {
			return new Parser97(keys);
		} else if (DocumentFactoryHelper.hasOOXMLHeader(stream)) {
			return new Parser07(keys);
		}
		throw new IllegalArgumentException("Your stream was neither an OLE2 stream, nor an OOXML stream.");
	}

	public static void main(String args[]) throws Exception {
		List<Key> keys = new ArrayList<>();
		{
			ParserProxy.parse(keys, new File("D:\\workspace\\Java\\Excalibur\\sample\\test.xlsx"));
			System.out.printf("%d\n", keys.size());
			for (Key key : keys) {
				System.out.printf("%s[%d] %s[%d]\n", key.getColumnName(), key.getColumnId(), key.getSheetName(), key.getSheetId());
			}
		} {
			ParserProxy.parse(keys, new File("D:\\workspace\\Java\\Excalibur\\sample\\test.xls"));
			System.out.printf("%d\n", keys.size());
			for (Key key : keys) {
				System.out.printf("%s[%d] %s[%d]\n", key.getColumnName(), key.getColumnId(), key.getSheetName(), key.getSheetId());
			}
		}
	}

}
