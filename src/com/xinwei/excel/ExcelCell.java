package com.xinwei.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelCell {
	private String columnName;
	private ExcelCellType type;
	private Cell cell;

	public ExcelCell() {
	}

	public ExcelCell(String columnName, Cell cell) {
		this(columnName, cell, ExcelCellType.STATIC);
	}

	public ExcelCell(String columnName, Cell cell, ExcelCellType type) {
		super();
		this.columnName = columnName;
		this.type = type;
		this.cell = cell;
	}

	public ExcelCellType getType() {
		return type;
	}

	public void setType(ExcelCellType type) {
		this.type = type;
	}

	public String getColumnName() {
		return columnName;
	}

	public void setColumnName(String columnName) {
		this.columnName = columnName;
	}

	public CellStyle getCellStyle() {
		return cell.getCellStyle();
	}

	public int getColumnIndex() {
		return cell.getColumnIndex();
	}

	public int getRowIndex() {
		return cell.getRowIndex();
	}

}
