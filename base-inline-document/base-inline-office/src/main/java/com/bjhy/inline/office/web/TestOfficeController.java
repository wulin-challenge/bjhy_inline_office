package com.bjhy.inline.office.web;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import javax.servlet.http.HttpServletRequest;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import com.bjhy.inline.office.base.FileConvert;
import com.bjhy.inline.office.base.OfficeStorePath;
import com.bjhy.inline.office.domain.Office;
import com.itextpdf.text.log.SysoCounter;

@Controller
@RequestMapping("/testOffice")
public class TestOfficeController {

	/**
	 * 打开office文档对应的html
	 * @param request
	 * @return
	 */
	@RequestMapping("testOpenOffice")
	public @ResponseBody Office openOffice(HttpServletRequest request){
//		String fileName = "aaa.doc"; 
//		String fileName = "bbb.docx"; 
//		String fileName = "dd.docx"; 
//		String fileName = "dc.doc";
//		String fileName = "xl1.xlsx";
//		String fileName = "xl1_1.xlsx";
//		String fileName = "xl3.xlsx";
//		String fileName = "xl4.xls";
//		String fileName = "xl5.xlsx";
//		String fileName = "xl6.xlsx";
//		String fileName = "aa.xls";
//		String fileName = "pp.ppt";
		String fileName = "pd.pdf";
//		String fileName = "tx.txt";
//		

		
		String[] fileNameArray = fileName.split("\\.");
		
		OfficeStorePath officeStorePath = new OfficeStorePath();
		FileInputStream stream = null;
		try {
			stream = new FileInputStream(new File(officeStorePath.getSourceDirectory(),fileName));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
		FileConvert fileConvert = new FileConvert(stream,fileNameArray[0],fileNameArray[1]);
//		FileConvert fileConvert = new FileConvert(stream,fileNameArray[0],fileNameArray[1],"pdf");
		Office office = fileConvert.getOffice();
		String html = office.getHtmlContext().replace("\n", " ");
		html = html.replace("&nbsp;", "   ");
		html = html.replace("<br>", "<br></br>");
		System.out.println(html);
//		officeStorePath.writeFile(fileStream, outPath);;
		return office;
	}
	
	/**
	 * 打开office文档对应的html
	 * @param request
	 * @return
	 */
	@RequestMapping("testOpenOffice2")
	public @ResponseBody Office openOffice2(String showType,String openFilePath,String fileName,String fileSuffix,HttpServletRequest request){
		
		if(StringUtils.isEmpty(openFilePath)){
			openFilePath = "";
		}
		OfficeStorePath officeStorePath = new OfficeStorePath();
		FileInputStream stream = null;
		try {
			stream = new FileInputStream(new File(officeStorePath.getSourceDirectory()+openFilePath,fileName+"."+fileSuffix));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
		FileConvert fileConvert = new FileConvert(stream,fileName,fileSuffix,showType);
		Office office = fileConvert.getOffice();
		return office;
	}
	
	//文件上传
	@RequestMapping(value = "fileUpload",method = RequestMethod.POST)
	public String fileUpload(MultipartFile uploadShowFile,HttpServletRequest request){
		String filePath = "";
		String fileName = "";
		try {
			if(uploadShowFile != null){
				fileName = uploadShowFile.getOriginalFilename();
				InputStream is = uploadShowFile.getInputStream();
				String showType = request.getParameter("showType");
				OfficeStorePath officeStorePath = new OfficeStorePath();
				
				filePath = officeStorePath.getUuid()+"/";
				FileUtils.copyInputStreamToFile(is, new File(officeStorePath.getSourceDirectory()+filePath+fileName));
				
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		request.setAttribute("data", "{'filePath':'"+filePath+"','fileName':'"+fileName+"'}");
		return "before/common/upload_result";
	}
	
	@RequestMapping("newPage")
	public String newPage(HttpServletRequest request){
		return "before/before_new_page";
	}
	
	@RequestMapping("newPage2")
	public String newPage2(HttpServletRequest request){
		return "before/before_new_page2";
	}

}
