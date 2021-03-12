package com.app.excel;

import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.ScriptStyle;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * jxl创建Excel文件示例
 */
public class JxlWriteDemo {
	public static void main(String[] args) throws IOException, RowsExceededException, WriteException {

		File xlsFile = new File("./src/main/resources/JXLMovieDemo.xls");
		// 创建工作簿
		WritableWorkbook workbook = Workbook.createWorkbook(xlsFile);
		// 创建一个工作表
		WritableSheet sheet = workbook.createSheet("电影表", 0);

		// 字体格式
		WritableFont font = new WritableFont(WritableFont.createFont("Microsoft YaHei UI"), // Excel对应字体名
				12, // 字号
				WritableFont.NO_BOLD, // 非粗体
				false, // 是否斜体
				UnderlineStyle.NO_UNDERLINE, // 无下划线
				Colour.BLACK, // 颜色
				ScriptStyle.NORMAL_SCRIPT// 脚本风格
		);

		// 单元格样式控制对象
		WritableCellFormat titleFormat = new WritableCellFormat(font);
		titleFormat.setAlignment(Alignment.CENTRE); // 设置单元格中的内容水平方向居中
		titleFormat.setBackground(Colour.LIGHT_GREEN); // 设置单元格的背景颜色
		titleFormat.setBorder(Border.ALL, BorderLineStyle.DOTTED);// 设置单元格边框
		titleFormat.setVerticalAlignment(VerticalAlignment.CENTRE); // 设置单元格中的内容垂直方向居中
		titleFormat.setWrap(false); // 是否自动换行
		titleFormat.setShrinkToFit(false); // 是否自动收缩（如果为true，则当内容过多时将自动缩小）

		for (int r = 0; r < 10; r++) {// 行
			for (int c = 0; c < 12; c++) {// 列
				// 向工作表中添加数据
				sheet.addCell(new Label(c, r, "星球大战" + (c + 1), titleFormat));
			}
		}

		for (int i = 0, numColumns = sheet.getColumns(); i < numColumns; i++) {
			sheet.setColumnView(i, 15);// 设置列宽
		}

		// 以Excel格式写出本工作簿中保存的数据
		workbook.write();
		// 关闭工作簿
		workbook.close();
	}
}
