package com.xinwei.report.month.activity;

import java.util.ArrayList;

public class ActivityMonth extends ArrayList<Activity> {
	private static final long serialVersionUID = -6829951462628401560L;
	
	// 项目月份
	private String month;

	public String getMonth() {
		return month;
	}

	public void setMonth(String month) {
		this.month = month;
	}

	@Override
	public String toString() {
		return "[项目月份=" + month + "," + super.toString() + "]";
	}

}
