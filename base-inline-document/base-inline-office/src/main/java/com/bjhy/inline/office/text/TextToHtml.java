package com.bjhy.inline.office.text;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

import org.apache.commons.io.IOUtils;

import com.bjhy.inline.office.base.OfficeStorePath;

/**
 * 将文件文件转换为在线预览的html
 * @author wubo
 *
 */
public class TextToHtml {

	/**
	 * 将txt转换为html
	 * @param storePath
	 * @return
	 */
	public String txtToHtml(OfficeStorePath storePath){
		
		StringBuffer html = new StringBuffer();
		String context = readFileContext(storePath);
		html.append("<div>"+context+"</div>");
		return html.toString();
	}
	
	/**
	 * 读取html的内容
	 */
	private String readFileContext(OfficeStorePath storePath){
		
		InputStream fileStream = storePath.getFileStream();
		
		BufferedInputStream bis = new BufferedInputStream(fileStream);
		String coder = OfficeStorePath.codeString(bis); 
		//html的文本内容
		StringBuffer htmlContext = OfficeStorePath.getText(bis, coder);
		String buf = htmlContext.toString();
			
		return buf;
	}
}
