package excel.separator;

import java.util.List;

public class CopyExcelSheet {

	public void createExcelSheetWithData(String excelFilePath, List<List<String>> dataList, String sheetName) {
		final WorkbookWriter workbookWriter = new WorkbookWriter(excelFilePath, dataList, sheetName);
		workbookWriter.writeSpreedSheet();
	}

	public static void main(String[] args) {

		CopyExcelSheet ces = new CopyExcelSheet();
		String sourcePath = "C:\\Users\\Maksim_Tolstik\\Documents\\temp\\largeFileSmall.xlsx";
		String destinationPath = "C:\\Users\\Maksim_Tolstik\\Documents\\temp\\largeFile_5.xlsx";

		final SheetExtractor sheetExtractor = new SheetExtractor(sourcePath, 0, 1300);
		List<List<String>> selectedRowDataList = sheetExtractor.getSheetDataList();
		String sheetName = sheetExtractor.getSheetName();

		ces.createExcelSheetWithData(destinationPath, selectedRowDataList, sheetName);

	}

}
