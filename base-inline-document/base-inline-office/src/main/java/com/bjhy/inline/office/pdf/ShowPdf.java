package com.bjhy.inline.office.pdf;

import java.io.File;
import java.io.IOException;

import com.bjhy.inline.office.base.OfficeStorePath;

/**
 * 页面一pdf的形式在页面上显示
 * @author wubo
 *
 */
public class ShowPdf {
	
	/**
	 * 将pdf转换为嵌入html中
	 * @param storePath
	 * @return
	 */
	public String pdfToHtml(OfficeStorePath storePath){

		StringBuffer html = new StringBuffer();
		//pdf存储路径
		String pdfStorePath = storePath.getImagesDirectory()+"/"+storePath.getFileName()+".pdf";
		
		//判断该文件是否存在,存在就必须跳过,否则就会造成流的内循环,类似于线程的死锁
		if(!new File(pdfStorePath).exists()){
			storePath.writeFile(storePath.getFileStream(), pdfStorePath);
		}else{
			try {
				storePath.getFileStream().close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		//pdf访问路径
		String pdfAccessPath = storePath.getURIResolverDirectory()+"/"+storePath.getFileName()+".pdf";
		
		html.append("<iframe width='98%' height='100%' src="+pdfAccessPath+"></iframe>");
		
		return html.toString();
	}

}
