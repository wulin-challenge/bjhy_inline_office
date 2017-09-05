package com.bjhy.inline.office.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

import com.bjhy.inline.office.domain.Office;
import com.bjhy.inline.office.excel.ExcelToHtml;
import com.bjhy.inline.office.pdf.HtmlToPdf;
import com.bjhy.inline.office.pdf.PdfToHtml;
import com.bjhy.inline.office.pdf.ShowPdf;
import com.bjhy.inline.office.ppt.Ppt2Html;
import com.bjhy.inline.office.text.TextToHtml;
import com.bjhy.inline.office.word.WordToHtml;

/**
 * 文件转换器
 * @author wubo
 *
 */
public class FileConvert {
	
	/**
	 * 文件输入流
	 */
	private InputStream fileInputSteam;
	
	private String fileName;
	
	/**
	 * 显示的类型
	 * pdf : 将文件转换为pdf
	 * html : 将文件转换为html
	 * doc : 将文件转换为word的doc类型
	 */
	private String showType = "html";
	
	/**
	 * 文件后缀
	 */
	private String suffix;
	
	/**
	 * office转换后的存储路径
	 */
	OfficeStorePath officeStorePath;
	
	/**
	 * 将私有的构造方法屏蔽掉
	 */
	@SuppressWarnings("unused")
	private FileConvert(){}

	public FileConvert(InputStream fileInputSteam,String fileName, String suffix) {
		this.fileInputSteam = fileInputSteam;
		this.fileName = fileName;
		this.suffix = suffix;
		officeStorePath = new OfficeStorePath(fileName,suffix);
	}
	
	public FileConvert(InputStream fileInputSteam,String fileName, String suffix,String showType) {
		this.fileInputSteam = fileInputSteam;
		this.fileName = fileName;
		this.suffix = suffix;
		this.showType = showType;
		officeStorePath = new OfficeStorePath(fileName,suffix);
	}
	
	public Office getOffice(){
		
		Office office = new Office();
		office.setFileName(fileName);
		office.setSuffix(suffix);
		
		if("docx".equalsIgnoreCase(suffix)){ //docx文件
			convertDocx(office);//转换office.后缀为docx的
			
		}else if("doc".equalsIgnoreCase(suffix)){
			convertDoc(office);//转换office.后缀为doc的
			
		}else if("xls".equalsIgnoreCase(suffix) || "xlsx".equalsIgnoreCase(suffix)){
			convertExcel(office);//转换office.后缀为xls与xlsx的
			
		}else if("ppt".equalsIgnoreCase(suffix) || "pptx".equalsIgnoreCase(suffix)){
			convertPpt(office);//转换office.后缀为ppt的
			
		}else if("pdf".equalsIgnoreCase(suffix) && "html".equalsIgnoreCase(showType)){
			convertPdf(office);//转换office.后缀为pdf的
			
		}else if("txt".equalsIgnoreCase(suffix)){
			convertTxt(office);//转换office.后缀为txt的
		}else{
			if(!"pdf".equalsIgnoreCase(suffix)){
				throw new RuntimeException("不支持后缀为  "+suffix+" 的文件");
			}
		}
		
		if("pdf".equalsIgnoreCase(showType)){
			
			//如果后缀不是pdf
			if(!"pdf".equalsIgnoreCase(suffix)){
				htmlToPdf(office);//将html转为pdf
			}
			
			htmlShowPdf(office);//在页面上以pdf的形式显示pdf
		}
		return office;
	}
	
	/**
	 * 将html转为pdf
	 * @param office
	 */
	private void htmlToPdf(Office office){
		try {
			HtmlToPdf htmlToPdf = new HtmlToPdf();
			htmlToPdf.htmlToPdf(office.getOfficeStorePath(), office);
			//修改文件的后缀
			suffix = "pdf";
			officeStorePath.setSuffix(suffix);
			fileInputSteam = new FileInputStream(new File(officeStorePath.getPdfStorePath()));
			officeStorePath.setFileStream(fileInputSteam);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}
	
	/**
	 * 转换office.后缀为docx的
	 */
	private void convertDocx(Office office){
		try {
			WordToHtml wordToHtml = new WordToHtml();
			
			officeStorePath.setFileStream(fileInputSteam);
			wordToHtml.docxToHtml(officeStorePath);
			office.setHtmlContext(officeStorePath.readHtmlContext());
			office.setOfficeStorePath(officeStorePath);
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}
	
	/**
	 * 转换office.后缀为doc的
	 */
	private void convertDoc(Office office){
		try {
			WordToHtml wordToHtml = new WordToHtml();
			
			officeStorePath.setFileStream(fileInputSteam);
			wordToHtml.docToHtml(officeStorePath);
			office.setHtmlContext(officeStorePath.readHtmlContext());
			office.setOfficeStorePath(officeStorePath);
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}
	
	/**
	 * 转换office.后缀为xls与xlsx的
	 * @throws InlineOfficeException 
	 */
	private void convertExcel(Office office){
		try {
			ExcelToHtml excelToHtml = new ExcelToHtml();
			
			officeStorePath.setFileStream(fileInputSteam);
			office.setHtmlContext(excelToHtml.readExcelToHtml(fileInputSteam, true));
			office.setOfficeStorePath(officeStorePath);
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}
	
	/**
	 * 转换office.后缀为ppt和pptx的
	 * @param office
	 */
	private void convertPpt(Office office){
		try {
			Ppt2Html ppt2Html = new Ppt2Html();
			
			officeStorePath.setFileStream(fileInputSteam);
			office.setHtmlContext(ppt2Html.pptToHtml(officeStorePath));
			office.setOfficeStorePath(officeStorePath);
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}
	
	/**
	 * 转换office.后缀为pdf的
	 * @param office
	 */
	private void convertPdf(Office office){
		try {
			PdfToHtml pdfToHtml = new PdfToHtml();
			officeStorePath.setFileStream(fileInputSteam);
			office.setHtmlContext(pdfToHtml.pdfToHtml(officeStorePath));
			office.setOfficeStorePath(officeStorePath);
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}
	
	/**
	 * 在页面上以pdf的形式显示pdf
	 * @param office
	 */
	private void htmlShowPdf(Office office){
		try {
			ShowPdf showPdf = new ShowPdf();
			officeStorePath.setFileStream(fileInputSteam);
			office.setHtmlContext(showPdf.pdfToHtml(officeStorePath));
			office.setOfficeStorePath(officeStorePath);
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}
	
	/**
	 * 转换office.后缀为txt的
	 * @param office
	 */
	private void convertTxt(Office office){
		try {
			TextToHtml textToHtml = new TextToHtml();
			officeStorePath.setFileStream(fileInputSteam);
			office.setHtmlContext(textToHtml.txtToHtml(officeStorePath));
			office.setOfficeStorePath(officeStorePath);
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}
}
