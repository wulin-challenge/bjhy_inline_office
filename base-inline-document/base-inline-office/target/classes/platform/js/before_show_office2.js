
$(function(){
	$("#openOfficeId").click(function(){
		var openFile = $("#openFile").val();
		setTimeout(callback(openFile,""),100);
	});
	
	$("#upload-show-file").change(function(){
		fileChangedealWith(); //change事件触发后处理的方法
		bandFileChangeEvent() //绑定file文件的change事件
	});
	
	
});

/**
 * 绑定file文件的change事件
 */
function bandFileChangeEvent(){
	$("#upload-show-form").empty();
	$("#upload-show-form").append("<a href='javascript:;' class='file'>上传显示 <input type='file' name='uploadShowFile' id='upload-show-file'></a>");
	$('#upload-show-file').off('change').on('change', function() {//上传
		fileChangedealWith(); //change事件触发后处理的方法
		bandFileChangeEvent() //绑定file文件的change事件
	　　return false;
	});
}

/**
 * change事件触发后处理的方法
 */
function fileChangedealWith(){
//	alert($('#upload-show-file').val());
	$("#upload-show-form").asyncSubmit(getUploadOption());
}

/**
 * 打开office文件
 * @param openFile
 * @param openFilePath
 */
function callback(openFile,openFilePath){
	return function(){
		var openWay = $("#openWay").val();
		openFile = openFile.split(".");
		
		var showType = $("#showType").val();
		var fileName = openFile[0];
		var fileSuffix = openFile[1];
		
		if('currentPage' == openWay){
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
		}else if('tabPage' == openWay){
			showType = encodeCode(showType);
			fileName = encodeCode(fileName);
			fileSuffix = encodeCode(fileSuffix);
			
			window.open(contextPath+'/testOffice/newPage2?showType='+showType+'&openFilePath='+openFilePath+'&fileName='+fileName+'&fileSuffix='+fileSuffix,'_blank');  
		}
	}
}

/**
 * 得到上传的option
 * @returns {___anonymous1852_1953}
 */
function getUploadOption(){
	var showType = $("#showType").val();
	var openWay = $("#openWay").val();
	
	var option = {
			check:function(){
				return true;
			},
			callback:function(data){
				setTimeout(callback(data.fileName,data.filePath),100);
			},
			data:{"showType":showType}
		};
	return option;
}

/**
 * 编码字符
 */
function encodeCode(chars){
	chars = window.encodeURI(chars);
	chars = window.encodeURI(chars);
	return chars;
}
