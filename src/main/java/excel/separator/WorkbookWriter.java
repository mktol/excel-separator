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

	private final int start;

	private final int finish;

	private List<String> headers;

	public WorkbookWriter(final String excelFilePath,
			final List<List<String>> dataList,
			final String sheetName,
			int start,
			int finish) {
		this.excelFilePath = excelFilePath;
		this.dataList = dataList;
		this.sheetName = sheetName;
		this.start = start;
		this.finish = finish;
		this.headers = dataList.get(0);
	}

	public void writeSpreedSheet() {
		validatePath(excelFilePath);
		try (Workbook workBook = new XSSFWorkbook();
				FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
			final Sheet newSheet = workBook.createSheet(sheetName);
			if (CollectionUtils.isNotEmpty(dataList)) {
				populateWorkBook(newSheet);
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

	private void populateWorkBook(final Sheet newSheet) {
		validateStartAndFinish();
		Row header = newSheet.createRow(0);
		populateRow(headers, header);
		int iter = 1;
		for (int i = start; i < finish; i++) {
			List<String> cellDataList = dataList.get(i);
			Row row = newSheet.createRow(iter++);
			populateRow(cellDataList, row);
		}
	}

	private void populateRow(final List<String> cellDataList, final Row row) {
		int columnNumber = cellDataList.size();

		for (int j = 0; j < columnNumber; j++) {
			String cellValue = cellDataList.get(j);
			row.createCell(j).setCellValue(cellValue);
		}
	}

	private void validateStartAndFinish() {
		final int size = dataList.size();
		if (!isStartAndFinishValid(size)) {
			throw new RuntimeException(
					"start or finish is not valid!" + "start =" + start + ", finish = " + finish + ", number of rows = "
							+ size);
		}
	}

	private boolean isStartAndFinishValid(final int size) {
		return start <= size && finish <= size && finish > 0 && start >= 0 && start < finish;
	}
}
