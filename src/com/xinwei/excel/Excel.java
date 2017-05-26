package com.xinwei.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 对excel进行操作工具类
 * 
 **/
public class Excel {

	protected Workbook workbook;

	protected Map<Sheet, HashMap<String, ExcelCell>> sheetTemplateMap = new HashMap<Sheet, HashMap<String, ExcelCell>>();

	public Excel() {
	}

	public Excel(Workbook templateWorkbook) {
		this.workbook = templateWorkbook;
	}

	public Excel(ExcelType excelType, int sheetNum) {
		this.workbook = createWorkbook(excelType, sheetNum);
	}

	public Excel(String tempFilePath, ExcelType excelType) throws Exception {
		this.workbook = getWorkbook(tempFilePath, excelType);
	}

	public Excel(InputStream is, ExcelType excelType) throws Exception {
		this.workbook = getWorkbook(is, excelType);
	}

	/**
	 * 资源关闭
	 * 
	 * @param targetFilePath
	 *            输出文件路径
	 */
	public void write(String targetFilePath) throws Exception {
		write(workbook, targetFilePath);
	}

	public void write(Workbook templateWorkbook, String targetFilePath) throws Exception {
		OutputStream os = null;
		try {
			File file = new File(targetFilePath);
			os = new FileOutputStream(file);
			write(templateWorkbook, os);
		} finally {
			if (os != null) {
				try {
					os.close();
				} catch (Exception e) {
				}
			}
		}
	}

	public void write(Workbook templateWorkbook, OutputStream os) throws Exception {
		templateWorkbook.write(os);
		os.flush();
	}

	public void close() throws Exception {
		close(workbook);
	}

	public void close(Workbook templateWorkbook) throws Exception {
		sheetTemplateMap.clear();
		templateWorkbook.close();
	}

	public Workbook getWorkbook(InputStream inputStream, ExcelType excelType) throws Exception {
		if (excelType == ExcelType.XLSX) {
			return new XSSFWorkbook(inputStream);
		} else if (excelType == ExcelType.XLS) {
			return new HSSFWorkbook(inputStream);
		}
		return null;
	}

	// 替换模板中的表头
	public void setTableHeader(int sheetIndex, Map<String, Object> dataMap) {
		setVariable(sheetIndex, dataMap);
	}

	// 替换模板中的內容
	public void setTableBody(int sheetIndex, List<Map<String, Object>> dataList) {
		setVariableList(sheetIndex, dataList);
	}

	/**
	 * 替换模板中的变量
	 */
	public void setVariable(int sheetIndex, Map<String, Object> dataMap) {
		Sheet sheet = getSheet(sheetIndex);
		setVariable(sheet, dataMap);
	}

	public void setVariable(Sheet sheet, Map<String, Object> dataMap) {
		if (dataMap == null || dataMap.isEmpty()) {
			return;
		}
		Set<Entry<String, Object>> entrySet = dataMap.entrySet();
		for (Entry<String, Object> entry : entrySet) {
			setVariable(sheet, entry.getKey(), entry.getValue());
		}
	}

	public void setVariable(int sheetIndex, String key, Object data) {
		Sheet sheet = getSheet(sheetIndex);
		setVariable(sheet, key, data);
	}

	public void setVariable(Sheet sheet, String key, Object data) {
		if (key == null) {
			return;
		}
		// 获取模板填充格式位置等数据
		HashMap<String, ExcelCell> template = getTemplate(sheet);

		// 获取对应单元格数据
		ExcelCell c = getCell(template, "${" + key + "}");

		setValue(data, sheet, c.getRowIndex(), c.getColumnIndex(), c.getCellStyle());
	}

	public void setVariableList(int sheetIndex, List<Map<String, Object>> dataList) {
		Sheet sheet = getSheet(sheetIndex);
		setVariableList(sheet, dataList);
	}

	public void setVariableList(Sheet sheet, List<Map<String, Object>> dataList) {
		if (dataList == null || dataList.size() == 0) {
			return;
		}
		// 获取模板填充格式位置等数据
		HashMap<String, ExcelCell> template = getTemplate(sheet);

		for (int i = 0; i < dataList.size(); i++) {
			Map<String, Object> dataMap = dataList.get(i);
			Set<Entry<String, Object>> entrySet = dataMap.entrySet();
			Row row = null;
			for (Entry<String, Object> entry : entrySet) {
				// 获取对应单元格数据
				ExcelCell c = getCell(template, "#{" + entry.getKey() + "}");
				if (row == null && i != 0) {
					row = insertRow(sheet, c.getRowIndex() + i);
				}
				// 写入数据
				setValue(entry.getValue(), sheet, c.getRowIndex() + i, c.getColumnIndex(), c.getCellStyle());
			}
		}
	}

	public void setDynamicVariable(int sheetIndex, String key, LinkedHashMap<String, CellStyle> dataMap) {
		Sheet sheet = getSheet(sheetIndex);
		setDynamicVariable(sheet, key, dataMap);
	}

	public void setDynamicVariable(Sheet sheet, String key, LinkedHashMap<String, CellStyle> columnNameMap) {
		if (key == null) {
			return;
		}
		// 获取模板填充格式位置等数据
		HashMap<String, ExcelCell> template = getTemplate(sheet);

		// 获取对应单元格数据
		ExcelCell c = getCell(template, "$${" + key + "}");
		Set<Entry<String, CellStyle>> entrySet = columnNameMap.entrySet();
		int i = 0;
		for (Entry<String, CellStyle> entry : entrySet) {
			String columnName = entry.getKey();
			CellStyle cellStyle = entry.getValue();
			setValue(columnName, sheet, c.getRowIndex(), c.getColumnIndex() + i++, cellStyle);
		}
	}

	public void setDynamicVariable(int sheetIndex, int row, int column, LinkedHashMap<String, CellStyle> columnNameMap) {
		Sheet sheet = getSheet(sheetIndex);
		setDynamicVariable(sheet, row, column, columnNameMap);
	}

	public void setDynamicVariable(Sheet sheet, int row, int column, LinkedHashMap<String, CellStyle> columnNameMap) {
		// 获取对应单元格数据
		Set<Entry<String, CellStyle>> entrySet = columnNameMap.entrySet();
		int i = 0;
		for (Entry<String, CellStyle> entry : entrySet) {
			String columnName = entry.getKey();
			CellStyle cellStyle = entry.getValue();
			setValue(columnName, sheet, row, column + i++, cellStyle);
		}
	}

	public void setDynamicVariableList(int sheetIndex, String key, LinkedHashMap<String, CellStyle> cellStyleMap, List<Map<String, Object>> dataList) {
		Sheet sheet = getSheet(sheetIndex);
		setDynamicVariableList(sheet, key, cellStyleMap, dataList);
	}

	public void setDynamicVariableList(Sheet sheet, String key, LinkedHashMap<String, CellStyle> cellStyleMap, List<Map<String, Object>> dataList) {
		if (key == null || dataList == null || dataList.size() == 0) {
			return;
		}
		// 获取模板填充格式位置等数据
		HashMap<String, ExcelCell> template = getTemplate(sheet);

		// 获取对应单元格数据
		ExcelCell c = getCell(template, "##{" + key + "}");

		for (int i = 0; i < dataList.size(); i++) {
			if (i != 0) {
				insertRow(sheet, c.getRowIndex() + i);
			}

			Map<String, Object> dataMap = dataList.get(i);

			Set<Entry<String, CellStyle>> entrySet = cellStyleMap.entrySet();
			int j = 0;

			for (Entry<String, CellStyle> entry : entrySet) {
				String columnName = entry.getKey();
				CellStyle cellStyle = entry.getValue();
				Object data = dataMap.get(columnName);
				setValue(data, sheet, c.getRowIndex() + i, c.getColumnIndex() + j++, cellStyle);
			}
		}
	}

	public void setDynamicVariableList(int sheetIndex, int row, int column, LinkedHashMap<String, CellStyle> cellStyleMap, List<Map<String, Object>> dataList) {
		Sheet sheet = getSheet(sheetIndex);
		setDynamicVariableList(sheet, row, column, cellStyleMap, dataList);
	}

	public void setDynamicVariableList(Sheet sheet, int row, int column, LinkedHashMap<String, CellStyle> cellStyleMap, List<Map<String, Object>> dataList) {
		if (dataList == null || dataList.size() == 0) {
			return;
		}

		for (int i = 0; i < dataList.size(); i++) {
			if (i != 0) {
				insertRow(sheet, row + i);
			}

			Map<String, Object> dataMap = dataList.get(i);

			Set<Entry<String, CellStyle>> entrySet = cellStyleMap.entrySet();
			int j = 0;
			for (Entry<String, CellStyle> entry : entrySet) {
				String columnName = entry.getKey();
				CellStyle cellStyle = entry.getValue();
				Object data = dataMap.get(columnName);
				setValue(data, sheet, row + i, column + j++, cellStyle);
			}
		}
	}

	/**
	 * 获取对应单元格样式等数据数据
	 * 
	 * @return
	 */
	private ExcelCell getCell(HashMap<String, ExcelCell> template, String key) {
		return template.get(key);
	}

	public Workbook createWorkbook(ExcelType excelType, int sheetNum) {
		Workbook workbook = null;
		if (excelType == ExcelType.XLSX) {
			workbook = new XSSFWorkbook();
		} else if (excelType == ExcelType.XLS) {
			workbook = new HSSFWorkbook();
		} else {
			workbook = new HSSFWorkbook();
		}
		for (int i = 0; i < sheetNum; i++) {
			workbook.createSheet();
		}
		return workbook;
	}

	public Workbook getWorkbook(String tempFilePath, ExcelType excelType) throws Exception {
		FileInputStream fileInputStream = null;
		try {
			fileInputStream = new FileInputStream(tempFilePath);
			return getWorkbook(fileInputStream, excelType);
		} catch (Exception e) {
			throw e;
		} finally {
			if (fileInputStream != null) {
				try {
					fileInputStream.close();
				} catch (Exception e) {
				}
			}
		}
	}

	public Row insertRow(Sheet sheet, Integer rowIndex) {
		Row row = null;
		if (sheet.getRow(rowIndex) != null) {
			int lastRowNo = sheet.getLastRowNum();
			sheet.shiftRows(rowIndex, lastRowNo, 1);// 将rowIndex到最后一行整体下移一行
		}
		row = sheet.createRow(rowIndex);
		return row;
	}

	public void deleteColumns(int sheetIndex, int... columns) {
		Sheet sheet = getSheet(sheetIndex);
		Iterator<Row> rowIter = sheet.iterator();
		while (rowIter.hasNext()) {
			Row row = rowIter.next();
			for (int column : columns) {
				Cell cell = row.getCell(column);
				row.removeCell(cell);
			}
		}
	}

	public void removeRows(int sheetIndex, int... rowIndexs) {
		Sheet sheet = getSheet(sheetIndex);
		for (int rowIndex : rowIndexs) {
			Row row = sheet.getRow(rowIndex);
			if (row != null) {
				sheet.removeRow(row);
			}
		}
	}

	public void shiftRows(int sheetIndex, int... rowIndexs) {
		Sheet sheet = getSheet(sheetIndex);
		for (int rowIndex : rowIndexs) {
			int lastRowNum = sheet.getLastRowNum();
			if (rowIndex == lastRowNum) {
				Row removingRow = sheet.getRow(rowIndex);
				if (removingRow != null) {
					sheet.removeRow(removingRow);
				}
			} else if (rowIndex >= 0 && rowIndex < lastRowNum) {
				sheet.shiftRows(rowIndex + 1, lastRowNum, -1);// 将行号为rowIndex+1一直到行号为lastRowNum的单元格全部上移一行，以便删除rowIndex行
			}
		}
	}

	public CellStyle createCellStyle() {
		if (workbook instanceof HSSFWorkbook) {
			return ((HSSFWorkbook) workbook).createCellStyle();
		} else if (workbook instanceof XSSFWorkbook) {
			return ((XSSFWorkbook) workbook).createCellStyle();
		}
		return null;
	}

	public Font createFont() {
		if (workbook instanceof HSSFWorkbook) {
			return ((HSSFWorkbook) workbook).createFont();
		} else if (workbook instanceof XSSFWorkbook) {
			return ((XSSFWorkbook) workbook).createFont();
		}
		return null;
	}

	public void setFont(int sheetIndex, int rowIndex, int columnIndex, Font font) {
		Cell cell = getCell(sheetIndex, rowIndex, columnIndex);

		cell.getCellStyle().setFont(font);
	}

	public void setFont(int sheetIndex, int rowIndex, int columnIndex, String fontName, short fontSize, boolean fontBold, boolean fontItalic) {
		Font font = createFont();
		font.setBold(fontBold);
		font.setFontName(fontName);
		font.setFontHeightInPoints(fontSize);
		font.setItalic(fontItalic);

		setFont(sheetIndex, rowIndex, columnIndex, font);
	}

	public void setFont(int sheetIndex, int rowIndex, int columnIndex, String fontName, short fontSize, boolean fontBold, boolean fontItalic, short fontColor) {
		Font font = createFont();
		font.setBold(fontBold);
		font.setFontName(fontName);
		font.setFontHeightInPoints(fontSize);
		font.setItalic(fontItalic);
		font.setColor(fontColor);
		setFont(sheetIndex, rowIndex, columnIndex, font);
	}

	public void setColumnWidth(int sheetIndex, int columnIndex, int width) {
		getSheet(sheetIndex).setColumnWidth(columnIndex, width);
	}

	public void setColumnHidden(int sheetIndex, int columnIndex, boolean hidden) {
		getSheet(sheetIndex).setColumnHidden(columnIndex, hidden);
	}

	public Cell getCell(int sheetIndex, int rowIndex, int columnIndex) {
		Sheet sheet = getSheet(sheetIndex);
		return getCell(sheet, rowIndex, columnIndex);
	}

	public Cell getCell(Sheet sheet, int rowIndex, int columnIndex) {
		Row row = sheet.getRow(rowIndex);
		return row.getCell(columnIndex);
	}

	public Cell getCell(Row row, int columnIndex) {
		return row.getCell(columnIndex);
	}

	public Cell findCell(int sheetIndex, String text) {
		return findCell(getSheet(sheetIndex), text);
	}

	public Cell findCell(Sheet sheet, String text) {
		int firstRowNum = sheet.getFirstRowNum();
		int lastRowNum = sheet.getLastRowNum();
		if (firstRowNum == -1 || lastRowNum == -1)
			return null;

		for (int rowIndex = firstRowNum; rowIndex <= lastRowNum; rowIndex++) {
			Cell cell = findCell(sheet, rowIndex, text);
			if (cell == null) {
				continue;
			} else {
				return cell;
			}
		}
		return null;
	}

	public Cell matchesCell(int sheetIndex, String text) {
		return matchesCell(getSheet(sheetIndex), text);
	}

	public Cell matchesCell(Sheet sheet, String text) {
		int firstRowNum = sheet.getFirstRowNum();
		int lastRowNum = sheet.getLastRowNum();
		if (firstRowNum == -1 || lastRowNum == -1)
			return null;

		for (int rowIndex = firstRowNum; rowIndex <= lastRowNum; rowIndex++) {
			Cell cell = matchesCell(sheet, rowIndex, text);
			if (cell == null) {
				continue;
			} else {
				return cell;
			}
		}
		return null;
	}

	public Cell findCell(int sheetIndex, int rowIndex, String text) {
		return findCell(getSheet(sheetIndex), rowIndex, text);
	}

	public Cell findCell(Sheet sheet, int rowIndex, String text) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			return null;
		}
		int firstCellNum = row.getFirstCellNum();
		int lastCellNum = row.getLastCellNum();
		if (firstCellNum == -1 || lastCellNum == -1) {
			return null;
		}

		for (int columnIndex = firstCellNum; columnIndex <= lastCellNum; columnIndex++) {
			Cell cell = findCell(row, columnIndex, text);
			if (cell == null) {
				continue;
			} else {
				return cell;
			}
		}
		return null;
	}

	public Cell matchesCell(int sheetIndex, int rowIndex, String text) {
		return matchesCell(getSheet(sheetIndex), rowIndex, text);
	}

	public Cell matchesCell(Sheet sheet, int rowIndex, String text) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			return null;
		}
		int firstCellNum = row.getFirstCellNum();
		int lastCellNum = row.getLastCellNum();
		if (firstCellNum == -1 || lastCellNum == -1) {
			return null;
		}

		for (int columnIndex = firstCellNum; columnIndex <= lastCellNum; columnIndex++) {
			Cell cell = matchesCell(row, columnIndex, text);
			if (cell == null) {
				continue;
			} else {
				return cell;
			}
		}
		return null;
	}

	public Cell findCell(Row row, int columnIndex, String text) {
		Cell cell = null;
		try {
			cell = getCell(row, columnIndex);
		} catch (Exception e) {
			e.printStackTrace();
		}
		if (cell == null) {
			return null;
		}

		String cellValue = cell.getStringCellValue();
		if (cellValue == null) {
			return null;
		}
		cellValue = cellValue.trim();
		if (cellValue.equals(text)) {
			return cell;
		}
		return null;
	}

	public Cell matchesCell(Row row, int columnIndex, String text) {
		Cell cell = null;
		try {
			cell = getCell(row, columnIndex);
		} catch (Exception e) {
			e.printStackTrace();
		}
		if (cell == null) {
			return null;
		}

		String cellValue = cell.getStringCellValue();
		if (cellValue == null) {
			return null;
		}
		cellValue = cellValue.trim();
		if (cellValue.matches(text)) {
			return cell;
		}
		return null;
	}

	public Cell findCell(Row row, String text) {
		for (Cell cell : row) {
			if (text.equals(getValue(cell))) {
				return cell;
			}
		}
		return null;
	}

	public Cell matchesCell(Row row, String text) {
		for (Cell cell : row) {
			if (getValue(cell).matches(text)) {
				return cell;
			}
		}
		return null;
	}

	public boolean isRow(Row row, String... texts) {
		for (String text : texts) {
			if (findCell(row, text) == null) {
				return false;
			}
		}
		return true;
	}

	public void setSheetName(int sheetIndex, String name) {
		setSheetName(workbook, sheetIndex, name);
	}

	public void setSheetName(Workbook templateWorkbook, int sheetIndex, String name) {
		if (templateWorkbook instanceof HSSFWorkbook) {
			((HSSFWorkbook) templateWorkbook).setSheetName(sheetIndex, name);
		} else if (templateWorkbook instanceof XSSFWorkbook) {
			((XSSFWorkbook) templateWorkbook).setSheetName(sheetIndex, name);
		}
	}

	public Sheet getSheet(int sheetIndex) {
		return workbook.getSheetAt(sheetIndex);
	}

	public int cloneSheet(int sheetIndex) {
		return cloneSheet(workbook, sheetIndex);
	}

	public int cloneSheet(Workbook templateWorkbook, int sheetIndex) {
		Sheet cloneSheet = templateWorkbook.cloneSheet(sheetIndex);
		return templateWorkbook.getSheetIndex(cloneSheet);
	}

	public void setValue(Object value, Sheet sheet, int rowIndex, int columnIndex, CellStyle cellStyle) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}
		Cell cell = row.getCell(columnIndex);
		if (cell == null) {
			cell = row.createCell(columnIndex);
		}
		if (cellStyle != null) {
			cell.setCellStyle(cellStyle);
		}
		// 对时间格式进行单独处理
		if (value == null) {
			cell.setCellValue("");
		} else {
			if (isCellDateFormatted(cellStyle)) {
				cell.setCellValue((Date) value);
			} else {
				cell.setCellValue(new XSSFRichTextString(value.toString()));
			}
		}
	}

	protected boolean isCellDateFormatted(CellStyle cellStyle) {
		if (cellStyle == null) {
			return false;
		}
		int i = cellStyle.getDataFormat();
		String f = cellStyle.getDataFormatString();

		return org.apache.poi.ss.usermodel.DateUtil.isADateFormat(i, f);
	}

	protected HashMap<String, ExcelCell> getTemplate(Sheet sheet) {
		HashMap<String, ExcelCell> template = sheetTemplateMap.get(sheet);
		if (template == null) {
			template = getSheetTemplate(sheet);
			sheetTemplateMap.put(sheet, template);
		}
		return template;
	}

	/**
	 * 读模板数据的样式值置等信息
	 * 
	 * @param keyMap
	 * @param sheet
	 */
	public HashMap<String, ExcelCell> getSheetTemplate(Sheet sheet) {
		HashMap<String, ExcelCell> keyMap = new HashMap<String, ExcelCell>();
		int firstRowNum = sheet.getFirstRowNum();
		int lastRowNum = sheet.getLastRowNum();
		if (firstRowNum == -1 || lastRowNum == -1)
			return keyMap;

		for (int rowIndex = firstRowNum; rowIndex <= lastRowNum; rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (row == null) {
				continue;
			}
			int firstCellNum = row.getFirstCellNum();
			int lastCellNum = row.getLastCellNum();
			if (firstCellNum == -1 || lastCellNum == -1) {
				continue;
			}

			for (int columnIndex = firstCellNum; columnIndex <= lastCellNum; columnIndex++) {
				Cell cell = row.getCell(columnIndex);
				if (cell == null) {
					continue;
				}

				String cellValue = getValue(cell);

				if (cellValue.startsWith("${") && cellValue.endsWith("}")) {
					String columnName = cellValue.substring(2, cellValue.length() - 1);
					keyMap.put(cellValue, new ExcelCell(columnName, cell, ExcelCellType.STATIC));
				} else if (cellValue.startsWith("#{") && cellValue.endsWith("}")) {
					String columnName = cellValue.substring(2, cellValue.length() - 1);
					keyMap.put(cellValue, new ExcelCell(columnName, cell, ExcelCellType.STATIC));
				} else if (cellValue.startsWith("$${") && cellValue.endsWith("}")) {
					String columnName = cellValue.substring(3, cellValue.length() - 1);
					keyMap.put(cellValue, new ExcelCell(columnName, cell, ExcelCellType.DYNAMIC));
				} else if (cellValue.startsWith("##{") && cellValue.endsWith("}")) {
					String columnName = cellValue.substring(3, cellValue.length() - 1);
					keyMap.put(cellValue, new ExcelCell(columnName, cell, ExcelCellType.DYNAMIC));
				}
			}
		}
		return keyMap;
	}

	public CellRangeAddress getMergedRegion(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return range;
				}
			}
		}
		return null;
	}

	public String getMergedRegionValue(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();

			if (row >= firstRow && row <= lastRow) {

				if (column >= firstColumn && column <= lastColumn) {
					Row fRow = sheet.getRow(firstRow);
					Cell fCell = fRow.getCell(firstColumn);
					return getValue(fCell);
				}
			}
		}

		return null;
	}

	public String getValue(int sheetIndex, int rowIndex, int columnIndex) {
		return getValue(getCell(getSheet(sheetIndex), rowIndex, columnIndex));
	}

	public String getValue(Sheet sheet, int rowIndex, int columnIndex) {
		return getValue(getCell(sheet, rowIndex, columnIndex));
	}

	public String getValue(Row row, int columnIndex) {
		return getValue(getCell(row, columnIndex));
	}

	public boolean isCellDateFormatted(Cell cell) {
		boolean flag = false;
		if (cell instanceof HSSFCell) {
			flag = HSSFDateUtil.isCellDateFormatted(cell);
		} else {
			flag = DateUtil.isCellDateFormatted(cell);
		}
		if (!flag) {
			short format = cell.getCellStyle().getDataFormat();
			if (format == 14 || format == 31 || format == 57 || format == 58 || format == 20 || format == 32) {
				flag = true;
			}
		}
		return flag;
	}

	public String getDateCellValue(Cell cell) {
		short format = cell.getCellStyle().getDataFormat();
		SimpleDateFormat sdf = null;
		if (format == 14 || format == 31 || format == 57 || format == 58) {
			// 日期
			sdf = new SimpleDateFormat("yyyy-MM-dd");
		} else if (format == 20 || format == 32) {
			// 时间
			sdf = new SimpleDateFormat("HH:mm");
		}
		double value = cell.getNumericCellValue();
		Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);
		return sdf.format(date);
	}

	/**
	 * 获取单元格的值
	 * 
	 * @param cell
	 * @return
	 */
	public String getValue(Cell cell) {

		if (cell == null)
			return "";

		if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

			return cell.getStringCellValue();

		} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {

			return String.valueOf(cell.getBooleanCellValue());

		} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

			return cell.getCellFormula();

		} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			if (isCellDateFormatted(cell)) {// 处理日期格式、时间格式
				return getDateCellValue(cell);
			} else {
				double value = cell.getNumericCellValue();
				CellStyle style = cell.getCellStyle();
				DecimalFormat format = new DecimalFormat();
				String temp = style.getDataFormatString();
				// 单元格设置成常规
				if ("General".equals(temp)) {
					format.applyPattern("#");
				}
				return format.format(value);
			}
		}
		return "";

	}
}