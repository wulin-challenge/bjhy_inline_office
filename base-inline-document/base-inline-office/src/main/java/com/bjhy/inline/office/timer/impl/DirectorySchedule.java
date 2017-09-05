package com.bjhy.inline.office.timer.impl;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.springframework.beans.factory.InitializingBean;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import com.bjhy.inline.office.base.OfficeStorePath;
import com.bjhy.inline.office.domain.CleanCatchFile;
import com.bjhy.inline.office.timer.Schedule;


/**
 * 清除目录定时器
 * @author wubo
 *
 */
@Component
public class DirectorySchedule implements Schedule,InitializingBean{
	
	@Override
	public void afterPropertiesSet() throws Exception {
	OfficeStorePath officeStorePath = new OfficeStorePath();
		
		String rootPath = officeStorePath.getRootPath();
		String htmlDirectory = rootPath+"/html/"; //html的临时目录
		String imagesDirectory = rootPath+"/images/"; //images的临时目录
		String sourceDirectory = officeStorePath.getSourceDirectory(); //source的临时目录
		
		officeStorePath.deleteFile(sourceDirectory);//删除source的临时目录
		officeStorePath.deleteFile(imagesDirectory);//删除html的临时目录
		officeStorePath.deleteFile(htmlDirectory);//删除images的临时目录
	}
	
	@Scheduled(cron = "50 * *  * * ? ")//每分钟在50秒时执行一次  
//	@Scheduled(cron = "0 0 1 * * ? ")//每一天的 1:00 AM
	public void timer() {
		OfficeStorePath officeStorePath = new OfficeStorePath();
		
		String rootPath = officeStorePath.getRootPath();
		String htmlDirectory = rootPath+"/html/"; //html的临时目录
		String imagesDirectory = rootPath+"/images/"; //images的临时目录
		String sourceDirectory = officeStorePath.getSourceDirectory(); //source的临时目录
		
		try {
			List<CleanCatchFile> cleanCatchFileList = getCleanCatch();//得到要清楚的缓存
			for (CleanCatchFile cleanCatchFile : cleanCatchFileList) {
				
				officeStorePath.deleteFile(sourceDirectory+cleanCatchFile.getCatchId());//删除source的临时目录
				officeStorePath.deleteFile(imagesDirectory+cleanCatchFile.getCatchId());//删除html的临时目录
				officeStorePath.deleteFile(htmlDirectory+cleanCatchFile.getCatchId());//删除images的临时目录
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	/**
	 * 得到要清楚的缓存
	 * @return
	 */
	private List<CleanCatchFile> getCleanCatch(){
		
		//要清除的缓存文件
		List<CleanCatchFile> cleanCatchFileList = new ArrayList<CleanCatchFile>();
		Map<String,CleanCatchFile> cleanCatchFileMap = OfficeStorePath.getCleanCatch();
		
		//克隆一份清除缓存
		Map<String,CleanCatchFile> cloneCleanCatch = cloneCleanCatch(cleanCatchFileMap);
		Set<Entry<String, CleanCatchFile>> entrySet = cloneCleanCatch.entrySet();
		for (Entry<String, CleanCatchFile> entry : entrySet) {
			
			if(isCleanCatch(entry.getValue())){//是否清除缓存文件
				cleanCatchFileMap.remove(entry.getKey());
				cleanCatchFileList.add(entry.getValue());
			}
		}
		
		return cleanCatchFileList;
	}
	
	/**
	 * 是否清除缓存文件
	 * @param cleanCatchFile
	 * @return
	 */
	private boolean isCleanCatch(CleanCatchFile cleanCatchFile){
		long catchTime = cleanCatchFile.getCatchTime().getTime();
		long currentTime = new Date().getTime();
		
		long difference = 2*60*1000; //清楚5分钟之前的临时文件 
		
		if((currentTime-catchTime)>=difference){
			return true;
		}
		return false;
	}
	
	/**
	 * 克隆一份清除缓存
	 */
	private Map<String,CleanCatchFile> cloneCleanCatch(Map<String,CleanCatchFile> cleanCatchFileMap){
		
		//克隆一份清除缓存
		Map<String,CleanCatchFile> cloneCleanCatch = new HashMap<String,CleanCatchFile>();
		
		Set<Entry<String, CleanCatchFile>> entrySet = cleanCatchFileMap.entrySet();
		for (Entry<String, CleanCatchFile> entry : entrySet) {
			
			CleanCatchFile cleanCatchFile = new CleanCatchFile();
			cleanCatchFile.setCatchId(entry.getValue().getCatchId());
			cleanCatchFile.setCatchTime(entry.getValue().getCatchTime());
			
			cloneCleanCatch.put(entry.getKey(),cleanCatchFile);
		}
		
		return cloneCleanCatch;
	}

}
