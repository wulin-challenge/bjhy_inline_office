package com.bjhy.inline.office.word;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLConnection;

import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.w3c.tidy.Configuration;
import org.w3c.tidy.Tidy;

public class HtmlToWord {
	
	public static void main(String[] args) {
		new HtmlToWord().writeWordFile();
	}
	
	public boolean writeWordFile() {    
        boolean w = false;    
        String path = "E:/";    
        try {    
            if (!"".equals(path)) {    
                // 检查目录是否存在    
                File fileDir = new File(path);    
                if (fileDir.exists()) {    
                    // 生成临时文件名称    
                    String fileName = "a.doc";                        
//            String content = gethtmlcode("http://homepage.yesky.com/59/2673059.shtml");  
              String content = getStringByInputStream(new FileInputStream(new File("D:/temp/show-web-word/word/4.html")));
            
//            org.jsoup.nodes.Document doc = Jsoup.parse(content);  
//            content=doc.html(); 
            System.out.println(content);
                    byte b[] = content.getBytes("utf-8");   
//                    content = new String(b,"utf-8");
                    
//                    XWPFDocument wordDoc = new XWPFDocument();
                    
                    ByteArrayInputStream bais = new ByteArrayInputStream(b);    
                    POIFSFileSystem poifs = new POIFSFileSystem();    
                    DirectoryEntry directory = poifs.getRoot();    
                    DocumentEntry documentEntry = directory.createDocument("WordDocument", bais);    
                    FileOutputStream ostream = new FileOutputStream(path+ fileName);    
                    poifs.writeFilesystem(ostream);  
//                    wordDoc.write(ostream);
                    bais.close();    
                    ostream.close();    
                }    
            }    
        } catch (IOException e) {    
            e.printStackTrace();    
      }    
      return w;    
    }    
	
	public static String gethtmlcode(String url){  
        StringBuffer str = new StringBuffer();  
        try {  
            URL u = new URL(url);  
            URLConnection uc = u.openConnection();  
            BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(uc.getInputStream(),"gbk"));
//            
            String buf;
            
            while((buf = bufferedReader.readLine()) != null){
            	str.append(buf);
            }
            
        }
        catch (IOException e) {  
            System.err.println(e);  
        }
  
        return str.toString();  
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
//			tidy.setXmlOut(true);
//			tidy.setXmlPIs(true);
//			tidy.setXmlPi(true);
//			tidy.setXmlSpace(true);
//			tidy.setXmlTags(true);
			 tidy.setXHTML(true); // 设定输出为xhtml(还可以输出为xml)  
			 tidy.setCharEncoding(Configuration.UTF8);
//         tidy.setIndentContent(true); // 缩进，可以省略，只是让格式看起来漂亮一些  
			 
			 tidy.setQuiet(true);                     
         tidy.setShowWarnings(false); //不显示警告信息  
			 tidy.setIndentContent(true);//  
			 tidy.setSmartIndent(true);  
         tidy.setIndentAttributes(false);  
//         tidy.setWraplen(1024); //多长换行  
//          tidy.setErrout(new PrintWriter(System.out));  
			 
			 tidy.parse(is, os);
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
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
    	BufferedReader bufferReader = new BufferedReader(new InputStreamReader(new ByteArrayInputStream(((ByteArrayOutputStream) os).toByteArray())));  
    	
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
				os.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
    	
    	return strBuffer.toString();
    }

}
