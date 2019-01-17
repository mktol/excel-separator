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

	Workbook excelWorkBook = null;

	private List<List<String>> sheetDataList = new ArrayList<>();

	private List<String> header = new ArrayList<>();
	private int firstRow;
	private int lastRow;

	public SheetExtractor(final String excelFilePath) {
		this.excelFilePath = excelFilePath;
		this.sheetDataList = extractSheetData();
	}

	private List<List<String>> extractSheetData() {
		validateFilePath();
		parseSheetData();
		return sheetDataList;
	}

	public List<List<String>> getSheetDataList() {
		return sheetDataList;
	}

	public String getSheetName() {
		return StringUtils.isEmpty(getSheet().getSheetName()) ? DEFAULT_SHEET_NAME : getSheet().getSheetName();
	}

	private void parseSheetData() {
		try (FileInputStream fis = new FileInputStream(excelFilePath.trim())) {

			excelWorkBook = new XSSFWorkbook(fis);
			Sheet copySheet = getSheet();
			firstRow = copySheet.getFirstRowNum();
			lastRow = copySheet.getLastRowNum();

			header = createListFromRow(copySheet, firstRow);
			/*  First row is excel file header, so read data from row next to it. */
			for (int i = firstRow; i < lastRow + 1; i++) {
				List<String> rowDataList = createListFromRow(copySheet, i);
				sheetDataList.add(rowDataList);
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

			putCellToListAsString(rowDataList, cell);

		}
		return rowDataList;
	}

	public int getNumberOfRows(){
		return lastRow - firstRow;
	}

	private void putCellToListAsString(final List<String> rowDataList, final Cell cell) {
		if (cell == null) {
			rowDataList.add(null);
			return;
		}
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

	public List<String> getHeader() {
		return new ArrayList<>(header);
	}
}
