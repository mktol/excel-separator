package excel.separator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class CopyExcelSheet {

	/* Get the data in excel file.
	 * Return: 2D String list contains specified row data.
	 * excelFilePath :  The exist file path need to copy.
	 * excelSheetName : Which sheet to copy.
	 * startRow : Start row number.
	 * endRow : End row number.
	 * */
	private List<List<String>> getExcelData(String excelFilePath, int startRow, int endRow) {

		List<List<String>> ret = new ArrayList<>();
		if (excelFilePath != null && !"".equals(excelFilePath.trim())) {
			try (FileInputStream fis = new FileInputStream(excelFilePath.trim())) {

				Sheet copySheet = getSheet(fis);

				int firstRow = copySheet.getFirstRowNum();
				int lastRow = copySheet.getLastRowNum();

				/*  First row is excel file header, so read data from row next to it. */
				for (int i = firstRow + 1; i < lastRow + 1; i++) {
					/* Only get desired row data. */
					if (satisfyRestriction(startRow, endRow, i)) {
						Row row = copySheet.getRow(i);

						int fCellNum = row.getFirstCellNum();
						int lCellNum = row.getLastCellNum();

						/* Loop in cells, add each cell value to the list.*/
						List<String> rowDataList = new ArrayList<>();
						for (int j = fCellNum; j < lCellNum; j++) {
							final Cell cell = row.getCell(j);
							if (cell == null) {
//								System.out.println(j);
								rowDataList.add(null);
								continue;
							}
							// Set the cell data value
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
							}

						}

						ret.add(rowDataList);
					}
				}

			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
		return ret;
	}

	private boolean satisfyRestriction(final int startRow, final int endRow, final int i) {
		return i >= startRow && i <= endRow;
	}

	private Sheet getSheet(final FileInputStream fis) throws IOException {
		Workbook excelWookBook = new XSSFWorkbook(fis);
		return excelWookBook.getSheetAt(0);
	}

	/* Create a new excel sheet with data.
	 * excelFilePath :  The exist excel file need to create new sheet.
	 * dataList : Contains all the data that need save to the new sheet.
	 * */
	private void createExcelSheetWithData(String excelFilePath, List<List<String>> dataList, String sheetName) {
		if (excelFilePath != null && !"".equals(excelFilePath.trim())) {
			try (Workbook workBook = new XSSFWorkbook()) {
				/* Open the exist file input stream. */
				final Sheet newSheet = workBook.createSheet(sheetName);
				/* Loop in the row data list, add each row data into the new sheet. */
				if (dataList != null) {
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

				/* Write the new sheet data to excel file */
				FileOutputStream fOut = new FileOutputStream(excelFilePath);
				workBook.write(fOut);
				fOut.close();

				System.out.println("New sheet added in excel file " + excelFilePath + " successfully. ");
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
	}

	public static void main(String[] args) {

		CopyExcelSheet ces = new CopyExcelSheet();
		String path = "C:\\Users\\Maksim_Tolstik\\Documents\\temp\\largeFileSmall.xlsx";
		String excelFilePath = "C:\\Users\\Maksim_Tolstik\\Documents\\temp\\largeFile_2.xlsx";

		String copySheetName = "Employee Info";

		List<List<String>> selectedRowDataList = ces.getExcelData(path, 0, 600);

		ces.createExcelSheetWithData(excelFilePath, selectedRowDataList, "asd");

	}

	private static class Pair {

		private String data;

		private CellType type;
	}
}
