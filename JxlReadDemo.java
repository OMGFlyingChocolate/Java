package com.app.excel;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * jxl读取Excel文件示例
 */
public class JxlReadDemo {
	public static void main(String[] args) throws IOException, BiffException {

		File xlsFile = new File("./src/main/resources/JXLMovieDemo.xls");
		// 获得工作簿对象
		Workbook workbook = Workbook.getWorkbook(xlsFile);
		// 获得所有工作表
		Sheet[] sheets = workbook.getSheets();

//		// 遍历工作表（方式一）
//		if (sheets != null) {
//			for (Sheet sheet : sheets) {
//				// 获得行数
//				int rows = sheet.getRows();
//				// 获得列数
//				int cols = sheet.getColumns();
//				// 读取数据
//				for (int row = 0; row < rows; row++) {
//					for (int col = 0; col < cols; col++) {
//						System.out.print(sheet.getCell(col, row).getContents() + "|");
//					}
//					System.out.println();
//				}
//			}
//		}

		// 遍历工作表（方式二）
		if (sheets != null) {
			for (Sheet sheet : sheets) {
				// 获得行数
				int rows = sheet.getRows();
				for (int row = 0; row < rows; row++) {// 行
					Cell[] cell = sheet.getRow(row);
					System.out.println("第【" + (row + 1) + "】行共【" + cell.length + "】列");
					for (int i = 0, length = cell.length; i < length; i++) {// 列
						switch (cell[i].getType().toString()) {
						case "Empty":
							System.out.print("Empty|");
							break;
						case "Label":
							System.out.print(cell[i].getContents() + "|");
							break;
						case "Number":
							break;
						case "Boolean":
							break;
						case "Error":
							break;
						case "Numerical Formula":
							break;
						case "Date Formula":
							break;
						case "String Formula":
							break;
						case "Boolean Formula":
							break;
						case "Formula Error":
							break;
						case "Date":
							break;
						}
					}
					System.out.println();
				}
			}
		}
		// 关闭工作簿对象
		workbook.close();
	}
}
