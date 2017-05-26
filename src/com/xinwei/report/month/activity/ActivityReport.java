package com.xinwei.report.month.activity;

import java.util.ArrayList;

public class ActivityReport extends ArrayList<ActivityMonth> {
	// 项目名称
	private String name;
	// 所在部门
	private String department;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getDepartment() {
		return department;
	}

	public void setDepartment(String department) {
		this.department = department;
	}

	@Override
	public String toString() {
		return "月度工作进展表 [项目名称=" + name + ", 所在部门=" + department + "," + super.toString() + "]";
	}

}
