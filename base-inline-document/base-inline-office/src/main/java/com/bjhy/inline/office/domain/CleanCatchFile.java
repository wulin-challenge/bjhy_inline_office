package com.bjhy.inline.office.domain;

import java.util.Date;

/**
 * 清楚缓存文件
 * @author wubo
 *
 */
public class CleanCatchFile {
	
	/**
	 * 缓存Id
	 */
	private String catchId;
	
	/**
	 * 缓存时间
	 */
	private Date catchTime;

	public String getCatchId() {
		return catchId;
	}

	public void setCatchId(String catchId) {
		this.catchId = catchId;
	}

	public Date getCatchTime() {
		return catchTime;
	}

	public void setCatchTime(Date catchTime) {
		this.catchTime = catchTime;
	}
	
}
