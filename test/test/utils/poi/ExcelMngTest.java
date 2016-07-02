package test.utils.poi;

import static org.junit.Assert.*;

import org.junit.Test;

import utils.poi.ExcelManager;

public class ExcelMngTest {

	public static void main (String[] args) {
		try {
			ExcelManager excelManager = ExcelManager.createExcelManager("D:\\pleiades\\workspace\\Utils\\tmp\\aaa.xlsx", "原本");

			System.out.println(excelManager);

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	@Test
	public void test() {

		try {
			ExcelManager excelManager = ExcelManager.createExcelManager("D:\\pleiades\\workspace\\Utils\\tmp\\POItest.xlsx");

		} catch (Exception e) {
			e.printStackTrace();
			fail();
		}
	}

}
