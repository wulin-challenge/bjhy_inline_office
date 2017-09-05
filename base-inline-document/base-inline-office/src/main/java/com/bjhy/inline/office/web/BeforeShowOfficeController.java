package com.bjhy.inline.office.web;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import javax.servlet.http.HttpServletRequest;

import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.bjhy.inline.office.base.FileConvert;
import com.bjhy.inline.office.base.OfficeStorePath;
import com.bjhy.inline.office.domain.Office;

/**
 * 显示Office
 * @author wubo
 */
@Controller
@RequestMapping("/beforeShowOffice")
public class BeforeShowOfficeController {
	
	@RequestMapping("index")
	public String index(HttpServletRequest request){
		return "before/before_show_office";
	}
	
	@RequestMapping("index2")
	public String index2(HttpServletRequest request){
		return "before/before_show_office2";
	}
	
	/**
	 * 打开office文档对应的html
	 * @param request
	 * @return
	 */
	@RequestMapping("openOffice")
	public @ResponseBody Office openOffice(HttpServletRequest request){
		String fileName = "dd.docx";
		
		String[] fileNameArray = fileName.split(".");
		
		OfficeStorePath officeStorePath = new OfficeStorePath();
		FileInputStream stream = null;
		try {
			stream = new FileInputStream(new File(officeStorePath.getSourceDirectory(),fileName));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
		FileConvert FileConvert = new FileConvert(stream,fileNameArray[0],fileNameArray[1]);
		return FileConvert.getOffice();
	}
	
	/**
	 * 在线修饰图片或者文件
	 * @return
	 */
	@RequestMapping(value="findOfficeImage")
	public ResponseEntity<byte[]> findOfficeImage(String imagePath){
		String suffix = imagePath.substring(imagePath.lastIndexOf(".")+1).trim();
		
		OfficeStorePath officeStorePath = new OfficeStorePath();
		String rootPath = officeStorePath.getRootPath();
		String reallyImagePath = rootPath+imagePath;
		
		FileInputStream imageInputStream = null;
		
		try {
			imageInputStream = new FileInputStream(new File(reallyImagePath));
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
		byte[] image = officeStorePath.getBytes(imageInputStream);
        HttpHeaders responseHeaders = new HttpHeaders();
        
        //在线显示pdf
        if("pdf".equalsIgnoreCase(suffix)){
        	  responseHeaders.setContentType(MediaType.parseMediaType("application/pdf"));
              responseHeaders.setContentLength(image.length);
              responseHeaders.set("Content-Disposition", "inline;filename=\""+System.currentTimeMillis()+".pdf\"");
        
        }else{
        	 responseHeaders.setContentType(MediaType.parseMediaType("image/png"));
             responseHeaders.setContentLength(image.length);
             responseHeaders.set("Content-Disposition", "attachment;filename=\""+System.currentTimeMillis()+".png\"");
        }
	        return new ResponseEntity<byte[]>(image,responseHeaders, HttpStatus.OK);
		}
	
}
