
$(function(){
	$("#open_word_html").click(function(){
//		window.open(contextPath+'/beforeShowOffice/openOffice','_blank');  
		
		$.ajax({
			   type: "get",
			   url:contextPath+'/testOffice/testOpenOffice',
//			   URL:CONTEXTPATH+'/BEFORESHOWOFFICE/OPENOFFICE',
			   dataType: "json",
			   async:true,  //async为true表示为异步请求,为false为同步请求
			   success: function(obj){
				   console.log(obj);
				   $("#office_id").empty();
				   $("#office_id").append(obj.htmlContext);
			   }
			}); 
	});
	
	$("#new_page_html").click(function(){
		var fileName = "xxxx文件名x";
			fileName = window.encodeURI(fileName);
			fileName = window.encodeURI(fileName);
		window.open(contextPath+'/testOffice/newPage?fileName='+fileName+'&fileSuffix=doc','_blank');  
	});
	
	
});