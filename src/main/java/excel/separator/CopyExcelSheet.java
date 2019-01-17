package excel.separator;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class CopyExcelSheet {

	private static int numberOfRowsPerFile = 5000;

	private static List<Section> calculateSections(final int numberOfRowsPerFile, final int totalNumberOfRows) {
		if (totalNumberOfRows <= numberOfRowsPerFile) {
			return Collections.singletonList(new Section(0, totalNumberOfRows));
		}

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

	public void createExcelSheetWithData(String destinationPath, List<List<String>> selectedRowDataList,
			String sheetName, Section section) {
		final WorkbookWriter workbookWriter = new WorkbookWriter(destinationPath,
				selectedRowDataList,
				sheetName,
				section.getStart(),
				section.getFinish());
		workbookWriter.writeSpreedSheet();
	}

	public static void main(String[] args) {

		CopyExcelSheet ces = new CopyExcelSheet();
		String sourcePath = "C:\\Users\\Maksim_Tolstik\\Documents\\temp\\largeFile.xlsx";
		String destinationPath = "C:\\Users\\Maksim_Tolstik\\Documents\\temp\\largeFile_7.xlsx";

		final SheetExtractor sheetExtractor = new SheetExtractor(sourcePath);
		List<List<String>> selectedRowDataList = sheetExtractor.getSheetDataList();
		String sheetName = sheetExtractor.getSheetName();

		final int totalNumberOfRows = sheetExtractor.getNumberOfRows();
		List<Section> sections = calculateSections(numberOfRowsPerFile, totalNumberOfRows);

		for (int i = 0; i < sections.size(); i++) {
			final Section section = sections.get(i);
			ces.createExcelSheetWithData("C:\\Users\\Maksim_Tolstik\\Documents\\temp\\small_file_" + i + ".xlsx",
					selectedRowDataList, sheetName, section);
		}
	}

}
