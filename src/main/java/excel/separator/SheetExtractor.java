package excel.separator;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class SheetExtractor {

	public static final String DEFAULT_SHEET_NAME = "Table 1";

	private final String excelFilePath;

	private final int startRow;

	private final int endRow;

	Workbook excelWorkBook = null;

	private List<List<String>> sheetDataList = new ArrayList<>();

	public SheetExtractor(final String excelFilePath, final int startRow, final int endRow) {
		this.excelFilePath = excelFilePath;
		this.startRow = startRow;
		this.endRow = endRow;
		this.sheetDataList = extractSheetData();
	}

	private List<List<String>> extractSheetData() {
		validateFilePath();
		populateSheetDataList();
		return sheetDataList;
	}

	public List<List<String>> getSheetDataList() {
		return sheetDataList;
	}

	public String getSheetName() {
		return StringUtils.isEmpty(getSheet().getSheetName()) ? DEFAULT_SHEET_NAME : getSheet().getSheetName();
	}

	private void populateSheetDataList() {
		try (FileInputStream fis = new FileInputStream(excelFilePath.trim())) {

			excelWorkBook = new XSSFWorkbook(fis);
			Sheet copySheet = getSheet();

			int firstRow = copySheet.getFirstRowNum();
			int lastRow = copySheet.getLastRowNum();

			/*  First row is excel file header, so read data from row next to it. */
			for (int i = firstRow + 1; i < lastRow + 1; i++) {
				/* Only get desired row data. */
				if (satisfyRestriction(startRow, endRow, i)) {
					List<String> rowDataList = createListFromRow(copySheet, i);
					sheetDataList.add(rowDataList);
				} else {
					break;
				}
			}

		} catch (Exception ex) {
			ex.printStackTrace();
			throw new RuntimeException(ex.getCause());
		}
	}

	private List<String> createListFromRow(final Sheet copySheet, final int i) {
		Row row = copySheet.getRow(i);

		int fCellNum = row.getFirstCellNum();
		int lCellNum = row.getLastCellNum();

		/* Loop in cells, add each cell value to the list.*/
		List<String> rowDataList = new ArrayList<>();
		for (int j = fCellNum; j < lCellNum; j++) {
			final Cell cell = row.getCell(j);
			if (cell == null) {
				rowDataList.add(null);
				continue;
			}
			// Set the cell data value
			putCellToListAsString(rowDataList, cell);

		}
		return rowDataList;
	}

	private void putCellToListAsString(final List<String> rowDataList, final Cell cell) {
		switch (cell.getCellType()) {
			case BLANK:
				rowDataList.add(cell.getStringCellValue());
				break;
			case BOOLEAN:
				rowDataList.add(Boolean.toString(cell.getBooleanCellValue()));
				break;
			case ERROR:
				rowDataList.add(Byte.toString(cell.getErrorCellValue()));
				break;
			case FORMULA:
				rowDataList.add(cell.getCellFormula());
				break;
			case NUMERIC:
				rowDataList.add(Double.toString(cell.getNumericCellValue()));
				break;
			case STRING:
				rowDataList.add(cell.getRichStringCellValue().getString());
				break;
			default:
				//do nothing
		}
	}

	private void validateFilePath() {
		if (!isFilePathNotEmpty()) {
			throw new RuntimeException("The file path is wrong or empty." + excelFilePath);
		}
	}

	private boolean isFilePathNotEmpty() {
		return excelFilePath != null && !"".equals(excelFilePath.trim());
	}

	private Sheet getSheet() {
		return excelWorkBook.getSheetAt(0);
	}

	private boolean satisfyRestriction(final int startRow, final int endRow, final int i) {
		return i >= startRow && i <= endRow;
	}
}
