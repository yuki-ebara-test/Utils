package utils.poi;

import java.util.ArrayList;
import java.util.List;

public class ExcelSheet {

	private String excelSheet;

	private List<ExcelRow> list = new ArrayList<ExcelRow>();
	/**
	 *
	 * @param excelSheet
	 */
	public ExcelSheet (String excelSheet) {
		this.excelSheet = excelSheet;
	}
	/**
	 *
	 * @param row
	 */
	public void addExcelRow (ExcelRow row) {
		list.add(row);
	}

	@Override
	public String toString() {
		return System.lineSeparator() + " " + excelSheet + list.toString() + System.lineSeparator();
	}



}
