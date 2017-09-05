$(function(){
	
	var showType = decodeCode(getUrlParam("showType"));
	var openFilePath = decodeCode(getUrlParam("openFilePath"));
	var fileName = decodeCode(getUrlParam("fileName"));
	var fileSuffix = decodeCode(getUrlParam("fileSuffix"));
	
	$.messager.progress({ 
	    title: '等待', 
	    msg: '文件打开中...', 
	    text: '后台正在打开文件,这可能需要一会儿....' 
	});
	
	PlatformUI.ajax({
		type: "get",
		   url:contextPath+'/testOffice/testOpenOffice2',
		   dataType: "json",
		   data:{"showType":showType,"openFilePath":openFilePath,"fileName":fileName,"fileSuffix":fileSuffix},
		   async:true,  //async为true表示为异步请求,为false为同步请求
		   afterOperation:function(obj){
			   $("#office_id").empty();
			   $("#office_id").append(obj.htmlContext);
			   $.messager.progress('close');
		   }
	});
});

/**
 * 解码字符
 * @param chars
 * @returns
 */
function decodeCode(chars){
	chars = window.decodeURI(chars);
	chars = window.decodeURI(chars);
	return chars;
}

/**
 * 得到当前跳转页面url的参数
 * @param name
 * @returns
 */
function getUrlParam(name){

	var reg = new RegExp("(^|&)"+ name +"=([^&]*)(&|$)"); 

	var r = window.location.search.substr(1).match(reg); 

	if (r!=null) return unescape(r[2]); return null;
} 