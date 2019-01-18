package excel.separator.util;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public final class FileUtils {

	private static final String EXTENSION_SEPARATOR = ".";

	private FileUtils() {
	}

	public static List<File> generateNewNames(String fileFullName, int numberOfNewFiles) {

		final String extension = fileFullName.substring(fileFullName.lastIndexOf(EXTENSION_SEPARATOR));
		final String firstPartOfName = fileFullName.substring(0, fileFullName.lastIndexOf(EXTENSION_SEPARATOR));
		validateExtension(extension);
		List<File> newFileNames = new ArrayList<>();
		for (int i = 0; i < numberOfNewFiles; i++) {
			String newName = firstPartOfName + "_" + i + extension;
			File newFile = new File(newName);
			newFileNames.add(newFile);
		}
		return newFileNames;
	}

	private static void validateExtension(final String extension) {
		if (!".xlsx".equals(extension)) {
			throw new RuntimeException("Not supported extendsion " + extension + "only .xlsx is spupported");

		}
	}
}
