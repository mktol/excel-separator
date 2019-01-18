package excel.separator;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class FileUtills {

	public static boolean isFilePresent(String filePath) {
		return new File(filePath).isFile();
	}

	public static boolean isFolderPresent(String filePath) {
		return new File(filePath).exists();
	}

	static List<File> generateNewNames(String fileFullName, int numberOfNewFiles) {

		List<File> newFileNames = new ArrayList<>();
		final String name = fileFullName;
		final String extension = name.substring(name.lastIndexOf("."));
		final String firlstPartOfName = name.substring(0, name.lastIndexOf("."));
		if (!".xlsx".equals(extension)) {
			throw new RuntimeException("Not supported extendsion " + extension + "only .xlsx is spupported");
		}
		for (int i = 0; i < numberOfNewFiles; i++) {
			String newName = firlstPartOfName + "_" + i + extension;
			File newFile = new File(newName);
			newFileNames.add(newFile);
		}
		return newFileNames;
	}

	private static String getFileExtension(File file) {
		String extension = "";

		try {
			if (file != null && file.exists()) {
				String name = file.getName();
				extension = name.substring(name.lastIndexOf("."));
			}
		} catch (Exception e) {
			extension = "";
		}

		return extension;

	}

}
