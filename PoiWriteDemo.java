package com.app.excel;

import java.io.File;
import java.io.IOException;

//import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * poi创建Excel文件示例
 */
public class PoiWriteDemo {
	public static void main(String[] args) throws IOException {

		// 创建工作薄
		HSSFWorkbook workbook = new HSSFWorkbook();
		// 创建一个新工作表
		HSSFSheet movie = workbook.createSheet("电影表");

//		HSSFRow row = movie.createRow(0);// 创建行
//		HSSFCell cell = row.createCell(0);// 创建单元格
//		cell.setCellValue("星球大战");

		for (int i = 0; i < 10; i++) {// 行
			HSSFRow row = movie.createRow(i);
			for (int j = 0; j < 12; j++) {// 列
				row.createCell(j).setCellValue("星球大战" + (j + 1));// 设置单元格的值
			}
		}

//		cell.setBlank();// 设置为空白单元格

		File xlsFile = new File("./src/main/resources/POIMovieDemo.xls");

		// 创建并写入文件
		workbook.write(xlsFile);
		// 关闭工作薄
		workbook.close();
	}
}
