package excel.separator;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.List;

public class WorkbookWriter {

	private final String excelFilePath;

	private final List<List<String>> dataList;

	private final String sheetName;

	public WorkbookWriter(final String excelFilePath, final List<List<String>> dataList, final String sheetName) {
		this.excelFilePath = excelFilePath;
		this.dataList = dataList;
		this.sheetName = sheetName;
	}

	public void writeSpreedSheet() {
		validatePath(excelFilePath);
		try (Workbook workBook = new XSSFWorkbook();
				FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
			final Sheet newSheet = workBook.createSheet(sheetName);
			if (CollectionUtils.isNotEmpty(dataList)) {
				populateWorkBook(dataList, newSheet);
			}
			workBook.write(fileOutputStream);
			System.out.println("New sheet added in excel file " + excelFilePath + " successfully. ");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	private void validatePath(final String excelFilePath) {
		if (StringUtils.isEmpty(excelFilePath)) {
			throw new RuntimeException("Path is empty");
		}
	}

	private void populateWorkBook(final List<List<String>> dataList, final Sheet newSheet) {
		int size = dataList.size();
		for (int i = 0; i < size; i++) {
			List<String> cellDataList = dataList.get(i);
			/* Create row to save the copied data. */
			Row row = newSheet.createRow(i + 1);

			int columnNumber = cellDataList.size();

			for (int j = 0; j < columnNumber; j++) {
				String cellValue = cellDataList.get(j);
				row.createCell(j).setCellValue(cellValue);
			}
		}
	}
}
