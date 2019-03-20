package com.cebbank.util;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("resource")
public class ExcelToTxt {

	public static String convertExcel(String excelPath) throws IOException {
		Workbook wb = null;
		Sheet sheet = null;
		Row row = null;
		List<Map<Integer, String>> list = null;
		String cellData = null;

		int x = excelPath.lastIndexOf(".");
		String textPath = excelPath.substring(0, x) + ".txt";
		wb = readExcel(excelPath);
		if (wb != null) {
			// 用来存放表中数据
			list = new ArrayList<Map<Integer, String>>();
			// 获取第一个sheet
			sheet = wb.getSheetAt(0);
			// 获取最大行数
			int rownum = sheet.getPhysicalNumberOfRows();
			// 获取第二行
			row = sheet.getRow(1);
			// 获取最大列数
			int colnum = row.getPhysicalNumberOfCells();
			for (int i = 0; i < rownum; i++) {
				Map<Integer, String> map = new LinkedHashMap<Integer, String>();
				row = sheet.getRow(i);
				if (row != null) {
					for (int j = 0; j < colnum; j++) {
						cellData = (String) getCellFormatValue(row.getCell(j));
						map.put(j, cellData);
					}
				} else {
					break;
				}
				list.add(map);
			}
		}
		// 遍历解析出来的list
		StringBuffer sb = new StringBuffer();
		for (int i = 0; i < list.size(); i++) {
			int k = list.get(i).entrySet().size();
			int j = 0;
			for (Entry<Integer, String> entry : list.get(i).entrySet()) {
				String value = entry.getValue();
				j++;
				if (j == k) {
					sb.append(value);
				} else {
					sb.append(value + "|");
				}

			}
			sb.append("\r\n");
		}
		writeToFile(sb.toString(), textPath);

		return textPath;
	}

	// 读取excel
	public static Workbook readExcel(String filePath) {
		Workbook wb = null;
		if (filePath == null) {
			return null;
		}
		String extString = filePath.substring(filePath.lastIndexOf(".")).toLowerCase();
		InputStream is = null;
		try {
			is = new FileInputStream(filePath);
			if (".xls".equals(extString)) {
				return wb = new HSSFWorkbook(is);
			} else if (".xlsx".equals(extString)) {
				return wb = new XSSFWorkbook(is);
			} else {
				return wb = null;
			}

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return wb;
	}

	public static Object getCellFormatValue(Cell cell) {
		Object cellValue = null;
		if (cell != null) {
			// 判断cell类型
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC: {
				cellValue = String.valueOf(cell.getNumericCellValue());
				break;
			}
			case Cell.CELL_TYPE_FORMULA: {
				// 判断cell是否为日期格式
				if (DateUtil.isCellDateFormatted(cell)) {
					// 转换为日期格式YYYY-mm-dd
					cellValue = cell.getDateCellValue();
				} else {
					// 数字
					cellValue = String.valueOf(cell.getNumericCellValue());
				}
				break;
			}
			case Cell.CELL_TYPE_STRING: {
				cellValue = cell.getRichStringCellValue().getString();
				break;
			}
			default:
				cellValue = "";
			}
		} else {
			cellValue = "";
		}
		return cellValue;
	}

	/**
	 * 生成文件
	 * 
	 * @param str
	 * @param filePath
	 * @throws IOException
	 */

	public static void writeToFile(String str, String filePath) throws IOException {
		BufferedWriter bw = null;
		try {
			FileOutputStream out = new FileOutputStream(filePath);// true,表示:文件追加内容，不重新生成,默认为false
			bw = new BufferedWriter(new OutputStreamWriter(out, "GBK"));
			bw.write(str += "\r\n");// 换行
			bw.flush();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			bw.close();
		}
	}
}
