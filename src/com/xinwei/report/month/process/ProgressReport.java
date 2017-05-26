package com.xinwei.report.month.process;

import java.util.ArrayList;
import java.util.List;

public class ProgressReport  {
	// 项目名称
	private String name;
	// 月度
	private String month;

	// 项目负责人
	private String charge;
	// 项目例会次数
	private String meetingNumber;
	// 项目其他管理
	private String otherManagement;
	// 项目重大问题
	private String majorIssues;
	// 评审次数
	private String reviewTimes;
	// 文档输出个数
	private String docOutputNumber;
	// 代码走读次数
	private String codeReviewTimes;
	// 版本输出个数
	private String versionOutputNumber;
	// 项目变更情况
	private String change;
	// 现场版本升级
	private String versionUpgrade;
	// 版本升级问题跟踪
	private String updateProblemTracking;
	// 改进计划
	private String improvementPlan;
	// 下月重大活动计划
	private String nextMonthPlan;

	
	private ArrayList<Progress> progressList=new ArrayList();
	
	
	
	public ArrayList<Progress> getProgressList() {
		return progressList;
	}

	public void setProgressList(ArrayList<Progress> progressList) {
		this.progressList = progressList;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getCharge() {
		return charge;
	}

	public void setCharge(String charge) {
		this.charge = charge;
	}

	public String getMeetingNumber() {
		return meetingNumber;
	}

	public void setMeetingNumber(String meetingNumber) {
		this.meetingNumber = meetingNumber;
	}

	public String getOtherManagement() {
		return otherManagement;
	}

	public void setOtherManagement(String otherManagement) {
		this.otherManagement = otherManagement;
	}

	public String getMajorIssues() {
		return majorIssues;
	}

	public void setMajorIssues(String majorIssues) {
		this.majorIssues = majorIssues;
	}

	public String getReviewTimes() {
		return reviewTimes;
	}

	public void setReviewTimes(String reviewTimes) {
		this.reviewTimes = reviewTimes;
	}

	public String getDocOutputNumber() {
		return docOutputNumber;
	}

	public void setDocOutputNumber(String docOutputNumber) {
		this.docOutputNumber = docOutputNumber;
	}

	public String getCodeReviewTimes() {
		return codeReviewTimes;
	}

	public void setCodeReviewTimes(String codeReviewTimes) {
		this.codeReviewTimes = codeReviewTimes;
	}

	public String getVersionOutputNumber() {
		return versionOutputNumber;
	}

	public void setVersionOutputNumber(String versionOutputNumber) {
		this.versionOutputNumber = versionOutputNumber;
	}

	public String getMonth() {
		return month;
	}

	public void setMonth(String month) {
		this.month = month;
	}

	public String getChange() {
		return change;
	}

	public void setChange(String change) {
		this.change = change;
	}

	public String getVersionUpgrade() {
		return versionUpgrade;
	}

	public void setVersionUpgrade(String versionUpgrade) {
		this.versionUpgrade = versionUpgrade;
	}

	public String getUpdateProblemTracking() {
		return updateProblemTracking;
	}

	public void setUpdateProblemTracking(String updateProblemTracking) {
		this.updateProblemTracking = updateProblemTracking;
	}

	public String getImprovementPlan() {
		return improvementPlan;
	}

	public void setImprovementPlan(String improvementPlan) {
		this.improvementPlan = improvementPlan;
	}

	public String getNextMonthPlan() {
		return nextMonthPlan;
	}

	public void setNextMonthPlan(String nextMonthPlan) {
		this.nextMonthPlan = nextMonthPlan;
	}

	@Override
	public String toString() {
		return "项目月度工作进展表 [项目名称=" + name + ", 月度=" + month + ", 项目负责人=" + charge + ", 项目例会次数=" + meetingNumber + ", 项目其他管理=" + otherManagement + ", 项目重大问题=" + majorIssues + ", 评审次数=" + reviewTimes
				+ ", 文档输出个数=" + docOutputNumber + ", 代码走读次数=" + codeReviewTimes + ", 版本输出个数=" + versionOutputNumber + ", 项目变更情况=" + change + ", 现场版本升级=" + versionUpgrade + ", 版本升级问题跟踪="
				+ updateProblemTracking + ", 改进计划=" + improvementPlan + ", 下月重大活动计划 =" + nextMonthPlan + "," + super.toString() + "]";
	}

}
