package com.xinwei.report.month;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

import com.xinwei.excel.Excel;
import com.xinwei.excel.ExcelType;
import com.xinwei.report.month.activity.ActivityMonth;
import com.xinwei.report.month.activity.ActivityReport;
import com.xinwei.report.month.process.Progress;
import com.xinwei.report.month.process.ProgressReport;

/**
 * 对excel进行操作工具类
 * 
 **/
public class MonthlyReportReader {

	@Test
	public void testMonthProcessReport() throws Exception {
		InputStream inputStream = this.getClass().getResourceAsStream("月报表.xls");

		System.out.println(getMonthlyProcessReport(inputStream, ExcelType.XLS));
	}

//	@Test
//	public void testMonthlyActivityReport() throws Exception {
//		InputStream inputStream = this.getClass().getResourceAsStream("9月报表0519.xls");
//
//		System.out.println(getMonthlyActivityReport(inputStream, ExcelType.XLS));
//	}
//
//	public List<ActivityReport> getMonthlyActivityReport(InputStream inputStream, ExcelType excelType) throws Exception {
//		Excel excel = new Excel(inputStream, excelType);
//		return getMonthlyActivityReport(excel);
//	}
//
//	public List<ActivityReport> getMonthActivityReport(String file, ExcelType excelType) throws Exception {
//		Excel excel = new Excel(file, excelType);
//		return getMonthlyActivityReport(excel);
//
//	}
//
//	protected List<ActivityReport> getMonthlyActivityReport(Excel excel) {
//		Sheet sheet = excel.getSheet(0);
//
//		String projectName = excel.getValue(sheet, 1, 0);
//		projectName = projectName.replace("项目名称：", "");
//		String department = excel.getValue(sheet, 1, 4);
//		department = department.replace("所在部门：", "");
//
//		ActivityReport activityReport = new ActivityReport();
//
//		activityReport.setName(projectName);
//		activityReport.setDepartment(department);
//
//		List<Row> activityMonthRowList = getActivityMonthRowList(excel, sheet);
//
//		for (int i = 0; i < activityMonthRowList.size(); i++) {
//			Row row = activityMonthRowList.get(i);
//
//			ActivityMonth activityMonth = new ActivityMonth();
//			activityMonth.setMonth(excel.getValue(row, 2));
//
//			int firstNum = row.getRowNum() + 3;// 内容行号
//
//			int lastNum = -1;
//
//			if (i == activityMonthRowList.size() - 1) {
//				lastNum = sheet.getLastRowNum();
//			} else {
//				lastNum = activityMonthRowList.get(i + 1).getRowNum();
//			}
//			if (lastNum == -1) {
//				break;
//			}
//			for (int j = firstNum; j < lastNum; j++) {
//				Row aRow = sheet.getRow(j);
//
//				// Activity activity = new Activity();
//				// activity.setSeq(excel.getValue(aRow, 0));
//				// activity.setTime(time);(excel.getValue(aRow, 1));
//				// activity.setSeq(excel.getValue(aRow, 2));
//				// activity.setSeq(excel.getValue(aRow, 3));
//				// activity.setSeq(excel.getValue(aRow, 4));
//				// activity.setSeq(excel.getValue(aRow, 5));
//				// activity.setSeq(excel.getValue(aRow, 6));
//				// activity.setSeq(excel.getValue(aRow, 7));
//				// activity.setSeq(excel.getValue(aRow, 8));
//				// activity.setSeq(excel.getValue(aRow, 9));
//
//			}
//
//		}
//
//		return null;
//	}
//
//	private List<Row> getActivityMonthRowList(Excel excel, Sheet sheet) {
//		int firstRowNum = sheet.getFirstRowNum();
//		int lastRowNum = sheet.getLastRowNum();
//		List<Row> rowList = new ArrayList<Row>();
//		for (int i = firstRowNum; i < lastRowNum - 1; i++) {
//			Row row = sheet.getRow(i);
//			boolean isRow = excel.isRow(row, "序号", "活动时间", "发现问题", "问题处理结果");
//			if (isRow) {
//				rowList.add(row);
//			}
//		}
//		return rowList;
//	}

	public List<ProgressReport> getMonthlyProcessReport(InputStream inputStream, ExcelType excelType) throws Exception {
		Excel excel = null;
		try {
			 excel = new Excel(inputStream, excelType);
			
			return getMonthlyProcessReport(excel);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			throw e;
		}
		finally
		{
			if(excel!=null)
			{
				excel.close();
			}
		}
	}

	public List<ProgressReport> getMonthlyProcessReport(String file, ExcelType excelType) throws Exception {
		Excel excel=null;
		try {
			 excel = new Excel(file, excelType);
			return getMonthlyProcessReport(excel);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			throw e;
		}
		finally
		{
			if(excel!=null)
			{
				excel.close();
			}
		}

	}

	protected List<ProgressReport> getMonthlyProcessReport(Excel excel) throws Exception {
		Sheet sheet = excel.getSheet(0);

		int startRowNum = 2;
		Row row = null;

		String month = excel.getValue(sheet, 0, 4);
		int sIndex = month.indexOf("（");
		int eIndex = month.indexOf("）");
		if ((sIndex != -1 || eIndex != -1) && sIndex < eIndex) {
			month = month.substring(sIndex + 1, eIndex).trim();
		}

		List<ProgressReport> reportList = new ArrayList<ProgressReport>();
		for (int rowIndex = startRowNum; rowIndex < sheet.getLastRowNum();) {
			row = sheet.getRow(rowIndex);
			ProgressReport progressReport = new ProgressReport();
            //按照位置读取相应的内容；
			progressReport.setMonth(month);
			progressReport.setName(excel.getValue(sheet, rowIndex, 0));
			progressReport.setCharge(excel.getValue(sheet, rowIndex, 1));
			progressReport.setMeetingNumber(excel.getValue(sheet, rowIndex, 2));
			progressReport.setOtherManagement(excel.getValue(sheet, rowIndex, 3));
			progressReport.setMajorIssues(excel.getValue(sheet, rowIndex, 9));
			progressReport.setReviewTimes(excel.getValue(sheet, rowIndex, 10));
			progressReport.setDocOutputNumber(excel.getValue(sheet, rowIndex, 11));
			progressReport.setCodeReviewTimes(excel.getValue(sheet, rowIndex, 12));
			progressReport.setVersionOutputNumber(excel.getValue(sheet, rowIndex, 13));
			progressReport.setChange(excel.getValue(sheet, rowIndex, 14));
			progressReport.setVersionUpgrade(excel.getValue(sheet, rowIndex, 15));
			progressReport.setUpdateProblemTracking(excel.getValue(sheet, rowIndex, 16));
			progressReport.setImprovementPlan(excel.getValue(sheet, rowIndex, 17));
			progressReport.setNextMonthPlan(excel.getValue(sheet, rowIndex, 18));

			int lastRow = rowIndex;
			CellRangeAddress mergedRegion = excel.getMergedRegion(sheet, rowIndex, 0);
			if (mergedRegion != null) {
				lastRow = mergedRegion.getLastRow();
			}

			for (int jobRow = rowIndex + 1; jobRow <= lastRow; jobRow++) {
				Progress progress = new Progress();
				progress.setSeq(excel.getValue(sheet, jobRow, 4));
				progress.setTime(excel.getValue(sheet, jobRow, 5));
				progress.setName(excel.getValue(sheet, jobRow, 6));
				progress.setContent(excel.getValue(sheet, jobRow, 7));
				progress.setOutput(excel.getValue(sheet, jobRow, 8));

				if (progress.getContent() != null && progress.getContent().trim().length() > 0) {
					progressReport.getProgressList().add(progress);
				}
			}

			rowIndex = lastRow + 1;

			if (progressReport.getName() != null && progressReport.getName().trim().length() > 0) {
				reportList.add(progressReport);
			}

		}
		excel.close();
		return reportList;
	}
}