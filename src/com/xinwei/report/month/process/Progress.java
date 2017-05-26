package com.xinwei.report.month.process;

public class Progress {
	// 序号
	private String seq;
	// 时间
	private String time;
	// 名称
	private String name;
	// 内容
	private String content;
	// 产出
	private String output;

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

	@Override
	public String toString() {
		return "主要工作内容 [序号=" + seq + ", 时间=" + time + ", 时间=" + name + ", 内容=" + content + ", 产出=" + output + "]";
	}

}
