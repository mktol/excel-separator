package excel.separator;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class SheetSeparator {

	public void executeSeparation(final int numberOfRowsPerFile, final String sourcePath,
			final String destinationPath) {

		final SheetExtractor sheetExtractor = new SheetExtractor(sourcePath);
		List<List<String>> selectedRowDataList = sheetExtractor.getSheetDataList();
		String sheetName = sheetExtractor.getSheetName();

		final int totalNumberOfRows = sheetExtractor.getNumberOfRows();
		List<Section> sections = calculateSections(numberOfRowsPerFile, totalNumberOfRows);

		final List<File> files = FileUtills.generateNewNames(sourcePath, sections.size());

		for (int i = 0; i < sections.size(); i++) {
			final Section section = sections.get(i);
			createExcelSheetWithData(files.get(i).getAbsolutePath(),
					selectedRowDataList, sheetName, section);
		}
	}

	public void createExcelSheetWithData(String destinationPath, List<List<String>> selectedRowDataList,
			String sheetName, Section section) {
		final WorkbookWriter workbookWriter = new WorkbookWriter(destinationPath,
				selectedRowDataList,
				sheetName,
				section.getStart(),
				section.getFinish());
		workbookWriter.writeSpreedSheet();
	}

	private List<Section> calculateSections(final int numberOfRowsPerFile, final int totalNumberOfRows) {
		if (totalNumberOfRows <= numberOfRowsPerFile) {
			return Collections.singletonList(new Section(0, totalNumberOfRows));
		}
		return getSections(numberOfRowsPerFile, totalNumberOfRows);

	}

	private List<Section> getSections(final int numberOfRowsPerFile, final int totalNumberOfRows) {
		List<Section> sections = new ArrayList<>();
		for (int i = 0; i <= totalNumberOfRows - numberOfRowsPerFile; i += numberOfRowsPerFile) {
			Section section = new Section(i, i + numberOfRowsPerFile);
			sections.add(section);
		}

		final int lastFileSize = totalNumberOfRows % numberOfRowsPerFile;

		if (lastFileSize != 0) {
			sections.add(new Section(totalNumberOfRows - lastFileSize, totalNumberOfRows));
		}
		return sections;
	}

	public static void main(String[] args) {

		System.out.println("\'"+args[1]+"\'");
		System.out.println("\'"+args[0]+"\'");
		int numberOfRowsPerFile = Integer.parseInt(args[1]);
		String sourcePath = args[0];

		new SheetSeparator().executeSeparation(numberOfRowsPerFile, sourcePath, "");
	}
}
