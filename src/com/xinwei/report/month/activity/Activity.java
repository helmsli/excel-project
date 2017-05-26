package com.xinwei.report.month.activity;

public class Activity {
	// 序号
	private String seq;
	// 活动时间
	private String time;
	// 活动名称
	private String name;
	// 活动内容
	private String content;
	// 活动产出
	private String output;
	// 评审次数
	private String reviewTimes;
	// 文档输出个数
	private String docOutputNumber;
	// 备注
	private String note;
	// 发现问题
	private String problems;
	// 创新点和亮点
	private String innovationAndhighlights;
	// 问题处理结果
	private String problemProcessingResult;
	// 是否一致
	private String isCconsistent;
	// 变化原因
	private String changeReason;
	// 改进完善
	private String improveAndPerfect;
	// 填表人
	private String formHolder;

	public String getSeq() {
		return seq;
	}

	public void setSeq(String seq) {
		this.seq = seq;
	}

	public String getTime() {
		return time;
	}

	public void setTime(String time) {
		this.time = time;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getContent() {
		return content;
	}

	public void setContent(String content) {
		this.content = content;
	}

	public String getOutput() {
		return output;
	}

	public void setOutput(String output) {
		this.output = output;
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

	public String getNote() {
		return note;
	}

	public void setNote(String note) {
		this.note = note;
	}

	public String getProblems() {
		return problems;
	}

	public void setProblems(String problems) {
		this.problems = problems;
	}

	public String getInnovationAndhighlights() {
		return innovationAndhighlights;
	}

	public void setInnovationAndhighlights(String innovationAndhighlights) {
		this.innovationAndhighlights = innovationAndhighlights;
	}

	public String getProblemProcessingResult() {
		return problemProcessingResult;
	}

	public void setProblemProcessingResult(String problemProcessingResult) {
		this.problemProcessingResult = problemProcessingResult;
	}

	public String getIsCconsistent() {
		return isCconsistent;
	}

	public void setIsCconsistent(String isCconsistent) {
		this.isCconsistent = isCconsistent;
	}

	public String getChangeReason() {
		return changeReason;
	}

	public void setChangeReason(String changeReason) {
		this.changeReason = changeReason;
	}

	public String getImproveAndPerfect() {
		return improveAndPerfect;
	}

	public void setImproveAndPerfect(String improveAndPerfect) {
		this.improveAndPerfect = improveAndPerfect;
	}

	public String getFormHolder() {
		return formHolder;
	}

	public void setFormHolder(String formHolder) {
		this.formHolder = formHolder;
	}

	@Override
	public String toString() {
		return "活动 [序号=" + seq + ", 活动时间=" + time + ", 活动名称=" + name + ", 活动内容=" + content + ", 活动产出=" + output + ", 评审次数=" + reviewTimes + ", 输出文档个数=" + docOutputNumber + ", 备注=" + note + ", 发现问题="
				+ problems + ", 创新点和亮点=" + innovationAndhighlights + ", 问题处理结果=" + problemProcessingResult + ", 是否一致=" + isCconsistent + ", 变化原因=" + changeReason + ", 改进完善=" + improveAndPerfect
				+ ", 填表人=" + formHolder + "]";
	}

}
