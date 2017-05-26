package com.xinwei.excel;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * 
 **/
public class ExcelWriter extends Excel {

	public ExcelWriter(ExcelType excelType, int sheetNum) {
		super(excelType, sheetNum);
	}

	public ExcelWriter(Workbook workbook) {
		super(workbook);
	}

	public ExcelWriter(String tempFilePath, ExcelType excelType) throws Exception {
		super(tempFilePath, excelType);
	}

	public ExcelWriter(InputStream is, ExcelType excelType) throws Exception {
		super(is, excelType);
	}

}