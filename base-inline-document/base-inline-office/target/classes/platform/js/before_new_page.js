$(function(){
	var fileName = getUrlParam("fileName");
	fileName = window.decodeURI(fileName);
	fileName = window.decodeURI(fileName);
	var fileSuffix = getUrlParam("fileSuffix");
	
	alert("fileName:"+fileName+"  ,  fileSuffix:"+fileSuffix);
	
	$.ajax({
		   type: "get",
		   url:contextPath+'/testOffice/testOpenOffice',
		   dataType: "json",
		   async:true,  //async为true表示为异步请求,为false为同步请求
		   success: function(obj){
			   console.log(obj);
			   $("#office_id").empty();
			   $("#office_id").append(obj.htmlContext);
		   }
		}); 
});

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