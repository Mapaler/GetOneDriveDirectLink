// JavaScript Document

//仿GM_getValue函数v1.0
if(typeof(GM_getValue) == "undefined")
{
	var GM_getValue = function(name, type){
		var value = localStorage.getItem(name);
		if (value == undefined) return value;
		if ((/^(?:true|false)$/i.test(value) && type == undefined) || type == "boolean")
		{
			if (/^true$/i.test(value))
				return true;
			else if (/^false$/i.test(value))
				return false;
			else
				return Boolean(value);
		}
		else if((/^\-?[\d\.]+$/i.test(value) && type == undefined) || type == "number")
			return Number(value);
		else
			return value;
	}
}
//仿GM_setValue函数v1.0
if(typeof(GM_setValue) == "undefined")
{
	var GM_setValue = function(name, value){
		localStorage.setItem(name, value);
	}
}
//仿GM_deleteValue函数v1.0
if(typeof(GM_deleteValue) == "undefined")
{
	var GM_deleteValue = function(name){
		localStorage.removeItem(name);
	}
}
//仿GM_listValues函数v1.0
if(typeof(GM_listValues) == "undefined")
{
	var GM_listValues = function(){
		var keys = new Array();
		for (var ki=0, kilen=localStorage.length; ki<kilen; ki++)
		{
			keys.push(localStorage.key(ki));
		}
		return keys;
	}
}

var maskObj = function(name,content) //一个掩码对象
{
	this.name = name;
	this.content = content;
	return this;
};
var masks = new Array(); //储存掩码数组
var curMask = 0; //当前选中的掩码

function mask_add()
{
	var mask_list = document.getElementsByClassName("mask-list")[0];
	var mask_name = document.getElementsByClassName("mask-name")[0];
	var mask_content = document.getElementsByClassName("mask-content")[0];
	if (mask_name.value.length>0 && mask_content.value.length>0)
	{
		addNewMask(mask_name.value,mask_content.value);
		mask_name.value = "";
		mask_content.value = "";
	}else
	{
		alert("掩码名或内容为空");
	}
	mask_list.selectedIndex = mask_list.options.length - 1;
	save_mask_local();
}
//从文本添加一个新的掩码
function addNewMask(name,content)
{
	var mask_list = document.getElementsByClassName("mask-list")[0];
	var mask = new maskObj(name,content);
	masks.push(mask);
	var opt = new Option(name + " : " + content, content);
	mask_list.options.add(opt);
}
function mask_remove()
{
	var mask_list = document.getElementsByClassName("mask-list")[0];
	var mask_name = document.getElementsByClassName("mask-name")[0];
	var mask_content = document.getElementsByClassName("mask-content")[0];
	if(mask_list.selectedIndex>=0)
	{
		masks.splice(mask_list.selectedIndex, 1);
		mask_list.remove(mask_list.selectedIndex);
	}else
	{
		alert("没有选中掩码");
	}
	save_mask_local();
}
function mask_select()
{
	var mask_list = document.getElementsByClassName("mask-list")[0];
	var mask_name = document.getElementsByClassName("mask-name")[0];
	var mask_content = document.getElementsByClassName("mask-content")[0];
	mask_name.value = masks[mask_list.selectedIndex].name;
	mask_content.value = masks[mask_list.selectedIndex].content;
	if (redata) generate_output(redata); //重新生成
	GM_setValue("godl-mask-index",mask_list.selectedIndex);
}
function save_mask_local() //把掩码设置保存到本地
{
	var mask_list = document.getElementsByClassName("mask-list")[0];
	var maskstr = JSON.stringify(masks);
	GM_setValue("godl-masks",maskstr);
	GM_setValue("godl-mask-index",mask_list.selectedIndex);
}
function load_mask_local() //从空白加载设置
{
	var mask_list = document.getElementsByClassName("mask-list")[0];
	var maskstr = GM_getValue("godl-masks");
	if (maskstr) //没有设置时就不保存了
	{
		var masksCfg = JSON.parse(maskstr);
		masksCfg.forEach(function(item){
			addNewMask(item.name,item.content);
		});
	}
	var mask_index = GM_getValue("godl-mask-index");
	mask_list.selectedIndex = mask_index;
}

function do_error(e)
{
	var outinfo = document.getElementsByClassName("outinfo")[0];
	var outcontent = document.getElementsByClassName("outcontent")[0];
	outinfo.innerHTML = "发生错误";
	outcontent.value = e.toString();
}
function do_cancel()
{
	var outinfo = document.getElementsByClassName("outinfo")[0];
	outinfo.innerHTML = "取消操作";
}
function do_success(files)
{
	redata = files;
	console.log(redata);
	generate_output(redata);
}

function generate_output(files)
{
	var mask_list = document.getElementsByClassName("mask-list")[0];
	var outinfo = document.getElementsByClassName("outinfo")[0];
	var outcontent = document.getElementsByClassName("outcontent")[0];
	var mask = masks[mask_list.selectedIndex];
	var filearr = files.value;
	
	outinfo.innerHTML = "共选择 " + filearr.length + " 个文件。"
	if (filearr.some(function(item){
		if (item.shared.scope != "anonymous") return true;
	})){
		outinfo.innerHTML += "存在非公共权限文件，注意添加通行许可代码。";
	}

	var outStrArr = filearr.map(function(item,index){
		var outStr = showMask(mask.content,item,index);
		return outStr;
	});
	outcontent.value = outStrArr.join("\n");
}

//显示掩码用
function showMask(str,file,index)
{
	var newTxt = str;
	var pattern = "%{([^}]+)}";
	var rs = null;
//	console.log(rs = regMask.exec(newTxt),rs = regMask.exec(newTxt),rs = regMask.exec(newTxt),rs = regMask.exec(newTxt))

	while (( rs = new RegExp(pattern).exec(newTxt) ) != null) {
        var mskO = rs[0], //包含括号的原始掩码
            mskN = rs[1]; //去掉掩码括号
		if (mskN != undefined)
		{
            mskN = (mskN != undefined) ? mskN.replace(/\\{/ig, "{").replace(/\\}/ig, "}").replace(/\\\\/ig, "\\") : null;
			try
			{
				var evTemp = eval("(" + mskN + ")");
				if (evTemp!=undefined)
					newTxt = newTxt.replace(mskO, evTemp.toString());
				else
					newTxt = newTxt.replace(mskO, "");
			}catch(e)
			{
				alert("掩码异常，详情查看控制台");
				console.error(mskO + " 掩码出现了异常情况",e);
			}
		}
	}

	return newTxt;
}

var redata;//储存返回的数据

window.onload = function() //网页加载初始化
{
	var localarr = GM_listValues();
	if (localarr.indexOf("godl-masks")<0) //没有掩码数据，初始化默认配置。
	{
		addNewMask("普通外链","http://storage.live.com/items/%{file.id}:/%{file.name}");
		addNewMask("最短链接","http://storage.live.com/items/%{file.id}");
		addNewMask("UBB代码外链图片","[img]http://storage.live.com/items/%{file.id}:/%{file.name}[/img]");
		addNewMask("掩码使用示例","在OneDrive里查看 %{file.name} 的地址是：%{file.webUrl}");
		addNewMask("掩码高级使用示例","%{index+1}号文件的尺寸是：%{file.size>1024?parseInt(file.size/1024)+\"KB\":file.size+\"B\"}");
	}
	load_mask_local();

	if (location.protocol !="https:" && location.hostname !="localhost" && location.hostname != "")
	{
		var goto = confirm("检测到你正在使用http模式，本应用要求使用https模式。\n是否自动跳转？");
		if (goto)
		{
			location.protocol = "https:";
		}
	}
}
//OneDrive官方API格式
function launchOneDrivePicker(){
	var outinfo = document.getElementsByClassName("outinfo")[0];
	outinfo.innerHTML = "正在等待API返回数据";
	var odOptions = {
		clientId: "d7b41a52-5bb3-43df-a20c-6259cb6a1886",
		action: "query", //share | download | query
		multiSelect: true,
		openInNewWindow: true,
		//advanced: {createLinkParameters: { type: "embed", scope: "anonymous" }},
		advanced: {queryParameters: "select=id,audio,createdBy,createdDateTime,cTag,deleted,description,eTag,file,fileSystemInfo,folder,image,lastModifiedBy,lastModifiedDateTime,location,name,package,parentReference,photo,remoteItem,searchResult,shared,sharepointIds,size,specialFolder,video,webDavUrl,webUrl"},
		success: function(files) {do_success(files); /* success handler */ },
		cancel: function() {do_cancel(); /* cancel handler */ },
		error: function(e) {do_error(e); /* error handler */ }
	};
	OneDrive.open(odOptions);
}