package com.bjhy.inline.office.base;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import org.apache.commons.io.FileUtils;
import org.w3c.tidy.Configuration;
import org.w3c.tidy.Tidy;

import com.bjhy.inline.office.domain.CleanCatchFile;

/**
 * office转换后的存储路径
 * @author wubo
 */
public class OfficeStorePath {
	
	/**
	 * 记录缓存文件
	 */
	private static Map<String,CleanCatchFile> cleanCatch = new HashMap<String,CleanCatchFile>(); 
	
	/**
	 * 一个OfficeStorePath对象只有一个uuid
	 */
	private final String uuid;
	
	private String fileName;//文件名称
	
	/**
	 * 文件后缀
	 */
	private String suffix;
	
	private InputStream fileStream;//文件流
	
	/**
	 * 当目录不存在时会创建基本的文件目录
	 */
	public OfficeStorePath(){
		this.uuid = UUID.randomUUID().toString().replace("-", "").toUpperCase();
		this.existOfDirectory();//目录是否存在
	}
	
	/**
	 * 判断 当目录不存在时是否会创建基本的文件目录
	 * @param isCreateDirectory 是否创建目录
	 */
	public OfficeStorePath(boolean isCreateDirectory){
		this.uuid = UUID.randomUUID().toString().replace("-", "").toUpperCase();
		if(isCreateDirectory){
			this.existOfDirectory();//目录是否存在
		}
	}
	
	/**
	 * 得到文件名和后缀,并创建基本的目录
	 * @param fileName 文件名称
	 * @param suffix 文件的后缀
	 */
	public OfficeStorePath(String fileName,String suffix){
		this.fileName = fileName;
		this.suffix = suffix;
		this.uuid = UUID.randomUUID().toString().replace("-", "").toUpperCase();
		this.existOfDirectory();//目录是否存在
	}
	
	/**
	 * 得到office转换后的根路径
	 * @return
	 */
	public String getRootPath(){
		File rootDir = new File(System.getProperty("user.dir"),"office");
		return rootDir.getAbsolutePath();
	}
	
	/**
	 * 得到office转换后存储html的目录
	 * @return
	 */
	private String getHtmlDirectory(){
		return getRootPath()+"/html/"+uuid+"/";
	}
	
	/**
	 * 得到Html文件路径
	 * @return
	 */
	public String getHtmlFilePath(){
		return getHtmlDirectory()+fileName+".html";
	}
	
	/**
	 * 得到图片目录
	 * @return
	 */
	public String getImagesDirectory(){
		return getRootPath()+"/images/"+uuid;
	}
	
	/**
	 * 得到URIResolver目录
	 * @return
	 */
	public String getURIResolverDirectory(){
		return "../beforeShowOffice/findOfficeImage?imagePath=/images/"+uuid;
	}
	
	/**
	 * 得到源文件目录
	 * @return
	 */
	public String getSourceDirectory(){
		return getRootPath()+"/source/";
	}
	
	/**
	 * 得到contextPath
	 * @return
	 */
	public String getContextPath(){
		String contextPath = InlineContext.contextPath;
		return contextPath;
	}
	
	/**
	 * 得到pdf的存放路径(得到pdf的存放文件)
	 * @return
	 */
	public String getPdfStorePath(){
		return getImagesDirectory()+"/"+getFileName()+".pdf";
	}
	
	/**
	 * 目录是否存在
	 */
	private void existOfDirectory(){
		
		//保存当前缓存信息
		CleanCatchFile cleanCatchFile = new CleanCatchFile();
		cleanCatchFile.setCatchId(uuid);
		cleanCatchFile.setCatchTime(new Date());
		cleanCatch.put(uuid,cleanCatchFile);
		
		File sourceDirectory = new File(getSourceDirectory());
		//源文件目录是否存在,不存在就创建
		if(!sourceDirectory.exists()){
			sourceDirectory.mkdirs();
		}
		
		File htmlDirectory = new File(getHtmlDirectory());
		
		//html文件夹目录是否存在,不存在就创建
		if(!htmlDirectory.exists()){
			htmlDirectory.mkdirs();
		}
		
		//存储图片路径目录是否存在,不存在就创建
		File imageDirectory = new File(getImagesDirectory());
		
		if(!imageDirectory.exists()){
			imageDirectory.mkdirs();
		}
		
		//存储字体路径目录是否存在,不存在就创建
		File fontDirectory = new File(getFontDirectory());
		
		if(!fontDirectory.exists()){
			fontDirectory.mkdirs();
		}
		
		//存储字体路文件是否存在,不存在就拷贝一份过来
		File fontPath = new File(getFontPath());
		if(!fontPath.exists()){
			InputStream simsunInputStream = Thread.currentThread().getContextClassLoader().getResourceAsStream("source/font/simsun.ttc");
			try {
				FileUtils.copyInputStreamToFile(simsunInputStream, new File(getFontPath()));
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		
		//默认的资源文件是否存在,不存在就copy一份过来
		isExistsSources();
	}
	
	/**
	 * 默认的资源文件是否存在,不存在就copy一份过来
	 */
	private void isExistsSources() {
		//源文件是否存在,不存在就拷贝一份过来
		List<String> defaultSource = getDefaultSourcePath();//默认资源存储的路径
		if(defaultSource != null && !defaultSource.isEmpty()){
			for (int i = 0; i < defaultSource.size(); i++) {
				
				String copySourcePath = getSourceDirectory()+defaultSource.get(i);
				File copySourceFile = new File(copySourcePath);
				if(!copySourceFile.exists()){
					//默认的存储的路径
					String basePath = "source/office-source/";
					InputStream sourceInputStream = Thread.currentThread().getContextClassLoader().getResourceAsStream(basePath+defaultSource.get(i));
					try {
						FileUtils.copyInputStreamToFile(sourceInputStream, new File(copySourcePath));
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
		}
	}
	
	/**
	 * 默认资源存储的路径
	 * @return
	 */
	private List<String> getDefaultSourcePath(){
		List<String> defaultSource = new ArrayList<String>();
		defaultSource.add("aa.xls");
		defaultSource.add("aaa.doc");
		defaultSource.add("cccc.docx");
		defaultSource.add("dd.docx");
		defaultSource.add("pp.ppt");
		defaultSource.add("pd.pdf");
		defaultSource.add("tx.txt");
		defaultSource.add("xl1.xlsx");
		defaultSource.add("xl3.xlsx");
		defaultSource.add("image1.jpg");
		defaultSource.add("image2.png");
		return defaultSource;
	}
	
	/**
	 * 删除某个文件夹及该文件夹下的所有文件及文件夹
	 * @param delpath
	 * @return
	 * @throws Exception
	 */
	public boolean deleteFile(String delpath) throws Exception{
		try {
			delpath = delpath.replace("\\\\", "/"); //// Java中4个反斜杠表示一个反斜杠
			delpath = delpath.replace("\\", "/"); 
			File file = new File(delpath);
			// 当且仅当此抽象路径名表示的文件存在且 是一个目录时，返回 true
			if (!file.isDirectory()) {
				file.delete();
			} else if (file.isDirectory()) {
				String[] filelist = file.list();
				for (int i = 0; i < filelist.length; i++) {
					File delfile = new File(delpath + "/" + filelist[i]);
					if (!delfile.isDirectory()) {
						delfile.delete();
						System.out.println(delfile.getAbsolutePath() + "删除文件成功");
					} else if (delfile.isDirectory()) {
						deleteFile(delpath + "/" + filelist[i]);
					}
				}
				System.out.println(file.getAbsolutePath() + "删除成功");
				file.delete();
			}

		} catch (FileNotFoundException e) {
			System.out.println("deletefile() Exception:" + e.getMessage());
		}
		return true;
	}
	
    /** 
     * 获得InputStream的byte数组 
     */  
	public byte[] getBytes(InputStream is) {
		byte[] buffer = null;
		ByteArrayOutputStream bos = new ByteArrayOutputStream(1000);
		byte[] b = new byte[1000];
		try {
			int n = 0;
			while ((n = is.read(b)) != -1) {
				bos.write(b, 0, n);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				is.close();
				bos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		buffer = bos.toByteArray(); 
		return buffer;
	}
	
	/**
	 * 读取html的内容
	 */
	public String readHtmlContext(){
		//html的文本内容
		StringBuffer htmlContext = new StringBuffer();
		try {
			File htmlFile = new File(getHtmlFilePath());
			
			//判断该html文件是否存在
			if(!htmlFile.exists()){
				System.out.println("该文件不存在 : "+htmlFile.getAbsolutePath());
				return null;
			}
			
			String context = FileUtils.readFileToString(htmlFile, "UTF-8");
			htmlContext.append(context);
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		return htmlContext.toString();
	}
	
	/**
	 * 向磁盘写文件
	 * @param fileStream
	 * @param outPath
	 */
	public void writeFile(InputStream fileStream,String outPath){
		FileOutputStream fileOutputStream = null;
		try {
			fileOutputStream = new FileOutputStream(new File(outPath));
			
			byte[] buf = new byte[1024];
			int n = 0;
			while((n=fileStream.read(buf)) != -1){
				fileOutputStream.write(buf,0,n);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			try {
				if(fileStream != null){
					fileStream.close();
				}
				
				if(fileOutputStream != null){
					fileOutputStream.close();
				}
				
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	  /**
     * 使用tidy来设置html文档
     * @param is
     * @return
     */
    public OutputStream getHtmlByTidy(InputStream is){
		OutputStream os = new ByteArrayOutputStream();
		try {
			Tidy tidy = new Tidy();
			tidy.setXHTML(true); // 设定输出为xhtml(还可以输出为xml)
			tidy.setCharEncoding(Configuration.UTF8);

			tidy.setQuiet(true);
			tidy.setShowWarnings(false); // 不显示警告信息
			tidy.setIndentContent(true);//
			tidy.setSmartIndent(true);
			tidy.setIndentAttributes(false);

			tidy.parse(is, os);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return os;
    }
    
    /**
     * 读出输入中的字符串
     * @param is
     * @return
     */
    public String getStringByInputStream(InputStream is){
    	BufferedReader bufferReader = new BufferedReader(new InputStreamReader(is));
    	
    	StringBuffer strBuffer = new StringBuffer();
    	String buf = null;
    	
    	try {
			while((buf = bufferReader.readLine()) != null){
				strBuffer.append(buf);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}finally{
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
    	
    	return strBuffer.toString();
    }
    
    /**
     * 将字符串变成输入流
     * @param str
     * @return
     */
    public InputStream getInputStreamFromString(String str){
    	ByteArrayInputStream byteArray = new ByteArrayInputStream(str.getBytes());
    	return byteArray;
    }
    
    /**
     * 读出输出中的字符串
     * @param is
     * @return
     */
    public String getStringByOutputStream(OutputStream os){
    	StringBuffer strBuffer = new StringBuffer();
    	try {
	    	BufferedReader bufferReader = new BufferedReader(new InputStreamReader(new ByteArrayInputStream(((ByteArrayOutputStream) os).toByteArray()),"UTF-8"));  
	    	
	    	String buf = null;
			while((buf = bufferReader.readLine()) != null){
				strBuffer.append(buf);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}finally{
			try {
				os.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
    	
    	return strBuffer.toString();
    }
    
    /**
     * 得到字体目录
     * @return
     */
    public String getFontDirectory(){
    	return getRootPath()+"/font/";
    }
    
    /**
     * 得到字体路径
     * @return
     */
    public String getFontPath(){
    	return getFontDirectory()+"simsun.ttc";
    }
    
	/**
	 * 得到文件名称
	 * @return
	 */
	public String getFileName() {
		return fileName;
	}

	/**
	 * 设置文件名称
	 * @param fileName
	 */
	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	/**
	 * 得到文件流
	 * @return
	 */
	public InputStream getFileStream() {
		return fileStream;
	}

	/**
	 * 设置文件流
	 * @param fileStream
	 */
	public void setFileStream(InputStream fileStream) {
		this.fileStream = fileStream;
	}

	/**
	 * 得到uuid
	 * @return
	 */
	public String getUuid() {
		return uuid;
	}


	public String getSuffix() {
		return suffix;
	}


	public void setSuffix(String suffix) {
		this.suffix = suffix;
	}

	public static Map<String,CleanCatchFile> getCleanCatch() {
		return cleanCatch;
	}
	
}
