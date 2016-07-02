package test.utils.poi;

import org.junit.Test;

import utils.poi.ExcelMng;

public class ExcelMngTest {

	public static void main (String[] args) {
		try {
			ExcelMng excelMng = ExcelMng.createExcelMng("D:\\pleiades\\workspace\\Utils\\tmp\\aaa.xlsx", "原本");

			System.out.println(excelMng);

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	@Test
	public void test() {

		try {
			ExcelMng excelMng = ExcelMng.createExcelMng("D:\\pleiades\\workspace\\Utils\\tmp\\POItest.xlsx");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
