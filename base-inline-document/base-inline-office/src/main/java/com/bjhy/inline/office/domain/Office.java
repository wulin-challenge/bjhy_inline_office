package com.bjhy.inline.office.domain;

import java.util.HashMap;
import java.util.Map;

import org.jsoup.Jsoup;

import com.bjhy.inline.office.base.OfficeStorePath;
import com.fasterxml.jackson.annotation.JsonIgnore;

public class Office {
	
	private String htmlContext;
	
	/**
	 * 文件名
	 */
	private String fileName;
	
	/**
	 * 文件后缀
	 */
	private String suffix;
	
	/**
	 * 额外参数
	 */
	private Map<String,Object> extraParams = new HashMap<String,Object>();
	
	@JsonIgnore
	private OfficeStorePath officeStorePath;

	public String getHtmlContext() {
		org.jsoup.nodes.Document doc = Jsoup.parse(htmlContext);  
		htmlContext=doc.html();  
		return htmlContext;
	}

	public void setHtmlContext(String htmlContext) {
		this.htmlContext = htmlContext;
	}

	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public String getSuffix() {
		return suffix;
	}

	public void setSuffix(String suffix) {
		this.suffix = suffix;
	}

	public Map<String, Object> getExtraParams() {
		return extraParams;
	}

	public void setExtraParams(Map<String, Object> extraParams) {
		this.extraParams = extraParams;
	}

	public OfficeStorePath getOfficeStorePath() {
		return officeStorePath;
	}

	public void setOfficeStorePath(OfficeStorePath officeStorePath) {
		this.officeStorePath = officeStorePath;
	}

}
