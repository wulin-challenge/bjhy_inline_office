package com.bjhy.inline.office.pdf;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.xhtmlrenderer.pdf.ITextFontResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import com.bjhy.inline.office.base.OfficeStorePath;
import com.bjhy.inline.office.domain.Office;
import com.lowagie.text.pdf.BaseFont;

/**
 * 将html转换为pdf
 * @author wubo
 *
 */
public class HtmlToPdf {
	
	/**
	 * html转换为pdf
	 * @param officeStorePath
	 * @param office
	 */
	public void htmlToPdf(OfficeStorePath officeStorePath,Office office){
		OutputStream os = null;
		//pdf的存储路径
		String pdfStorePath = officeStorePath.getPdfStorePath();
		 try {
	        	os = new FileOutputStream(pdfStorePath);
				ITextRenderer renderer = new ITextRenderer();  
				String html = office.getHtmlContext();
				
				if(StringUtils.isEmpty(html)){
					html = "";
				}
				
				String pageSize = getPdfPageWidthAndHeight(office);
				
				String fontFamily = "<head><style type='text/css' mce_bogus='1'>body {font-family: SimSun} "+pageSize+" </style>";
				html = html.replace("\0","");
				html = html.replace("font-family","xxxx");
				html = html.replace("font-weight","xxxx");
//				html = html.replace("font-size","xxxx");
				html = html.replace("<head>",fontFamily);
				
				office.setHtmlContext(html);
				//得到xhtml格式的html
//				html = officeStorePath.getStringByOutputStream(officeStorePath.getHtmlByTidy(officeStorePath.getInputStreamFromString(html)));
				html = officeStorePath.getStringByOutputStream(officeStorePath.getHtmlByTidy(IOUtils.toInputStream(html, "UTF-8")));
				
				renderer.setDocumentFromString(html);
				// 解决中文支持问题  
				
				ITextFontResolver fontResolver = renderer.getFontResolver();  
				fontResolver.addFont(officeStorePath.getFontPath(), BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);  
				//解决图片的相对路径问题  
				String imageUrl = officeStorePath.getContextPath()+officeStorePath.getURIResolverDirectory().substring(2);
//				imageUrl = "/analysis-sjzfjc/beforeShowOffice/findOfficeImage?imagePath=/images/DFF2316ABA2445678B97CDC9ECCAA467";
				renderer.getSharedContext().setBaseURL(imageUrl); 
				renderer.layout();
				renderer.createPDF(os);
			} catch (Exception e) {
				e.printStackTrace();
			} finally{
				try {
					os.flush();
					os.close();  
				} catch (IOException e) {
					e.printStackTrace();
				}  
			}
	}
	
	/**
	 * 根据不同的文件得到pdf不同的大小
	 * //这是A4纸的大小 : @page{size:210mm 297mm;}
	 * //这是A3纸的大小 : @page{size:297mm 420mm;}
	 * //这是A2纸的大小 : @page{size:420mm 594mm;}
	 * @param office
	 * @return
	 */
	private String getPdfPageWidthAndHeight(Office office){
		String suffix = office.getSuffix();
		String pageSize = "@page{size:490mm 594mm;}";
		
		if("docx".equalsIgnoreCase(suffix)){ //docx文件
			pageSize = "@page{size:420mm 594mm;}";
			
		}else if("doc".equalsIgnoreCase(suffix)){
			pageSize = "@page{size:420mm 594mm;}";
			
		}else if("xls".equalsIgnoreCase(suffix) || "xlsx".equalsIgnoreCase(suffix)){
			pageSize = "@page{size:500mm 594mm;}";
			
		}else if("ppt".equalsIgnoreCase(suffix) || "pptx".equalsIgnoreCase(suffix)){
			pageSize = "@page{size:297mm 420mm;}";
			
		}else if("pdf".equalsIgnoreCase(suffix)){
			pageSize = "@page{size:420mm 594mm;}";
			
		}else if("txt".equalsIgnoreCase(suffix)){
			pageSize = "@page{size:500mm 594mm;}";
		}
		return pageSize;
	}
}
