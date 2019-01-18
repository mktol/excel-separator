package excel.separator;

import excel.separator.core.SheetSeparator;

public class Main {

	public static void main(String[] args) {

		int numberOfRowsPerFile = Integer.parseInt(args[1]);
		String sourcePath = args[0];

		new SheetSeparator().executeSeparation(numberOfRowsPerFile, sourcePath, "");
	}
}
