package com.bjhy.inline.office.text;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

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
		//html的文本内容
		StringBuffer htmlContext = new StringBuffer();
		BufferedReader fileReader = null;
		try {
			InputStreamReader inputStreamReader = new InputStreamReader(storePath.getFileStream(),"UTF-8");
			fileReader = new BufferedReader(inputStreamReader);
			String buf = null;
			buf = fileReader.readLine();
			while(buf != null){
				
				//特殊字符转义
				buf = buf.replace("\"", "&quot;");
				buf = buf.replace("&", "&amp;");
				buf = buf.replace("<", "&lt;");
				buf = buf.replace(">", "&gt;");
				
				htmlContext.append(buf+"<br>");
				buf = fileReader.readLine();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			if(fileReader != null){
				try {
					fileReader.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return htmlContext.toString();
	}
}
