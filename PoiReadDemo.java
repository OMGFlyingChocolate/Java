package com.app.excel;

import java.io.File;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * poi读取Excel文件示例
 */
public class PoiReadDemo {
	public static void main(String[] args) throws IOException {

		File xlsFile = new File("./src/main/resources/POIMovieDemo.xls");
		// 获取工作簿
		Workbook workbook = WorkbookFactory.create(xlsFile);
		// 获得工作表个数
		int sheetCount = workbook.getNumberOfSheets();
		// 遍历工作表
		for (int i = 0; i < sheetCount; i++) {
			// 根据索引获取对应的工作簿操作对象
			Sheet sheet = workbook.getSheetAt(i);

			System.out.println("开始遍历工作簿【" + sheet.getSheetName() + "】");

			// 获得行数
			int rowNumber = sheet.getLastRowNum() + 1;

			System.out.println("该工作簿有【" + rowNumber + "】行");

			for (int rownum = 0; rownum < rowNumber; rownum++) {// 遍历行
				Row row = sheet.getRow(rownum);

				short minColIx = row.getFirstCellNum();
				short maxColIx = row.getLastCellNum();

//				System.out.println("minColIx:" + minColIx + "|maxColIx:" + maxColIx);

				for (short colIx = minColIx; colIx < maxColIx; colIx++) {// 遍历列
					Cell cell = row.getCell(colIx);

//					if (cell == null) {
//						continue;
//					}
//					System.out.printf("%10s", cell.getStringCellValue());

					String strVal = cell == null ? "EMPTY" : cell.getStringCellValue();
					System.out.printf("%10s", strVal);
				}
				System.out.println();
			}
		}
		// 关闭工作簿
		workbook.close();
	}
}
