document.writeln("<table width=100% height=100% cellpadding=2 cellspacing=0><tr><td align=left><select onChange=\"FormatText('FormatBlock',this[this.selectedIndex].value);\"><option selected>段落</option><option value=\"&lt;P&gt;\">正文</option><option value=\"&lt;H1&gt;\">标题一</option><option value=\"&lt;H2&gt;\">标题二</option><option value=\"&lt;H3&gt;\">标题三</option><option value=\"&lt;H4&gt;\">标题四</option><option value=\"&lt;H5&gt;\">标题五</option><option value=\"&lt;H6&gt;\">标题六</option><option value=\"&lt;PRE&gt;\">预设格式</option></select> <select name=\"selectFont\" onChange=\"FormatText('fontname', selectFont.options[selectFont.selectedIndex].value);\"><option selected>字体<option value=\"removeFormat\">默认字体<option value=\"宋体\">宋体<option value=\"黑体\">黑体<option value=\"隶书\">隶书<option value=\"幼圆\">幼圆<option value=\"楷体_GB2312\">楷体<option value=\"仿宋_GB2312\">仿宋<option value=\"新宋体\">新宋体<option value=\"华文彩云\">华文彩云<option value=\"华文仿宋\">华文仿宋<option value=\"华文新魏\">华文新魏<option value=\"Arial\">Arial<option value=\"Arial Black\">Arial Black<option value=\"Arial Narrow\">Arial Narrow<option value=\"Century\">Century<option value=\"Courier New\">Courier New<option value=\"Georgia\">Georgia<option value=\"Impact\">Impact<option value=\"Lucida Console\">Lucida Console<option value=\"MS Sans Serif\">MS Sans Serif<option value=\"System\">System<option value=\"Symbol\">Symbol<option value=\"Tahoma\">Tahoma<option value=\"Times New Roman\">Times New Roman&nbsp; &nbsp;<option value=\"Verdana\">Verdana<option value=\"Webdings\">Webdings<option value=\"Wingdings\">Wingdings</option></select> <select onChange=\"FormatText('fontsize',this[this.selectedIndex].value);\" name=\"D2\"><option class=\"heading\" selected>字体大小<option value=1>一号<option value=2>二号<option value=3>三号<option value=4>四号<option value=5>五号<option value=6>六号<option value=7>七号</option></select>");
document.writeln("<img src=images/ybb/fbcolor.gif alt=\"突出颜色\" align=absmiddle style=cursor:hand onClick=BackColor()>");
document.writeln("<img src=images/ybb/fgcolor.gif alt=\"字体颜色\" align=absmiddle style=cursor:hand onClick=foreColor()>");
document.writeln("<img src=images/ybb/replace.gif alt=\"替换\" align=absmiddle style=cursor:hand onClick=replace()>");
document.writeln("<img src=images/ybb/img.gif alt=\"插入图片\" align=absmiddle style=cursor:hand onClick=img()>");
document.writeln("<img alt=\"插入FLASH文件\" src=images/ybb/swf.gif align=absmiddle style=cursor:hand onclick=flash()>");
document.writeln("<img alt=\"插入Media Player文件\" src=images/ybb/mp.gif align=absmiddle style=cursor:hand onclick=MediaPlayer()>");
document.writeln("<img alt=\"插入RealPlayer文件\" src=images/ybb/rm.gif align=absmiddle style=cursor:hand onclick=RealPlay()>");
document.writeln("<img alt=\"插入表情\" src=images/ybb/em.gif align=absmiddle style=cursor:hand onclick=em()>");

var FormatTextlist="插入超链接 createLink|去掉超链接 Unlink|<br>|粗体 bold|倾斜 italic|下划线 underline|上标 superscript|下标 subscript|删除线 strikethrough|删除文字格式 RemoveFormat|左对齐 Justifyleft|居中 JustifyCenter|右对齐 JustifyRight|两端对齐 justifyfull|编号 insertorderedlist|项目符号 InsertUnorderedList|减少缩进量 Outdent|增加缩进量 indent|普通水平线 InsertHorizontalRule|剪切 cut|复制 copy|粘贴 paste|撤消 undo|恢复 redo|全选 selectAll|取消选择 unselect|删除当前选中区 Delete"
var list= FormatTextlist.split ('|'); 
for(i=0;i<list.length;i++) {
if (list[i]=="<br>"){document.write("<br>");
}else{
var TextName= list[i].split (' '); 
document.write(" <img align=absmiddle src=images/ybb/"+TextName[1]+".gif alt="+TextName[0]+" style=cursor:hand onClick=FormatText('"+TextName[1]+"')> ");
}
}
document.writeln("<img alt=\"清理代码\" src=images/ybb/CleanCode.gif align=absmiddle style=cursor:hand onclick=CleanCode()>");
document.writeln("</td></tr><tr><td height=100%><iframe ID=HtmlEditor MARGINHEIGHT=5 MARGINWIDTH=5 width=100% height=100%></iframe></td></tr><tr></tr></table>");

if (document.all){var IframeID=frames["HtmlEditor"];}else{var IframeID=document.getElementById("HtmlEditor").contentWindow;}
if (navigator.appVersion.indexOf("MSIE 6.0",0)==-1){IframeID.document.designMode="On"}
IframeID.document.open();
IframeID.document.write ('<script>i=0;function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13 && i==0){i=1;parent.document.yuziform.content.value=document.body.innerHTML;parent.document.yuziform.submit();parent.document.yuziform.submit1.disabled=true;}}<\/script><body onkeydown=ctlent()>');
IframeID.document.close();
IframeID.document.body.contentEditable = "True";
IframeID.document.body.innerHTML=document.getElementById("content").value;
IframeID.document.body.style.fontSize="10pt";


function em(){
var arr = showModalDialog("inc/Emotion.htm", "", "dialogWidth:20em; dialogHeight:9.5em; status:0;help:0");
if (arr != null){
IframeID.focus()
sel=IframeID.document.selection.createRange();
sel.pasteHTML(arr);
}
}

function CleanCode(){
var body = IframeID.document.body;
var html = IframeID.document.body.innerHTML;
html = html.replace(/\<p>/gi,"[$p]");
html = html.replace(/\<\/p>/gi,"[$\/p]");
html = html.replace(/\<br>/gi,"[$br]");
html = html.replace(/\<[^>]*>/g,"");
html = html.replace(/\[\$p\]/gi,"<p>");
html = html.replace(/\[\$\/p\]/gi,"<\/p>");
html = html.replace(/\[\$br\]/gi,"<br>");
IframeID.document.body.innerHTML = html;
}


function FormatText(command,option){IframeID.focus();IframeID.document.execCommand(command,true,option);}

function CheckLength(){alert("最大字符为 "+50000+ " 字节\n您的内容已有 "+IframeID.document.body.innerHTML.length+" 字节");}

function emoticon(theSmilie){
IframeID.focus();
sel=IframeID.document.selection.createRange();
sel.pasteHTML("<img src=images/Emotions/"+theSmilie+".gif>");
}

function DoTitle(addTitle) {
var revisedTitle;var currentTitle = document.yuziform.topic.value;revisedTitle = addTitle+currentTitle;document.yuziform.topic.value=revisedTitle;document.yuziform.topic.focus();
}

function CheckForm(form){
form.content.value=IframeID.document.body.innerHTML; 
MessageLength=IframeID.document.body.innerHTML.length;
if(MessageLength<2){alert("文章内容不能小于2个字符！");return false;}
if(MessageLength>50000){alert("文章内容不能超过50000个字符！");return false;}
document.yuziform.submit1.disabled = true;
}

function Gopreview()
{
document.preview.content.value=IframeID.document.body.innerHTML; 
window.open('', 'preview_page', 'resizable,scrollbars,width=750,height=450');
document.preview.submit()
}

function BackColor()
{
  var arr = showModalDialog("inc/selcolor.htm", "", "dialogWidth:18em; dialogHeight:17.5em; status:0;help:0");
  if (arr != null) FormatText('BackColor', arr);
  else IframeID.focus();
}

function foreColor()
{
  var arr = showModalDialog("inc/selcolor.htm", "", "dialogWidth:18em; dialogHeight:17.5em; status:0;help:0");
  if (arr != null) FormatText('forecolor', arr);
  else IframeID.focus();
}

//////替换内容
function replace()
{
  var arr = showModalDialog("inc/replace.html", "", "dialogWidth:22em;dialogHeight:10em;status:0;help:0");
	if (arr != null){
		var ss;
		ss = arr.split("*")
		a = ss[0];
		b = ss[1];
		i = ss[2];
		con = IframeID.document.body.innerHTML;
		if (i == 1)
		{
			con = bbsxp_rCode(con,a,b,true);
		}else{
			con = bbsxp_rCode(con,a,b);
		}
		IframeID.document.body.innerHTML = con;
	}
	else IframeID.focus();
}
function bbsxp_rCode(s,a,b,i){
	a = a.replace("?","\\?");
	if (i==null)
	{
		var r = new RegExp(a,"gi");
	}else if (i) {
		var r = new RegExp(a,"g");
	}
	else{
		var r = new RegExp(a,"gi");
	}
	return s.replace(r,b); 
}
//////替换内容结束

function img(){
url=prompt("请输入图片文件地址:","http://");
if(!url || url=="http://") return;
IframeID.focus();
sel=IframeID.document.selection.createRange();
sel.pasteHTML("<img src="+url+">");
}

function RealPlay(){
var arr = showModalDialog("inc/RealPlay.htm", "", "dialogWidth:22em; dialogHeight:10.5em; status:0;help:0");
if (arr != null){
IframeID.focus()
sel=IframeID.document.selection.createRange();
sel.pasteHTML(arr);
}
}

function MediaPlayer(){
var arr = showModalDialog("inc/MediaPlayer.htm", "", "dialogWidth:22em; dialogHeight:10.5em; status:0;help:0");
if (arr != null){
IframeID.focus()
sel=IframeID.document.selection.createRange();
sel.pasteHTML(arr);
}
}

function flash(){
var arr = showModalDialog("inc/flash.htm", "", "dialogWidth:22em; dialogHeight:9em; status:0;help:0");
if (arr != null){
IframeID.focus()
sel=IframeID.document.selection.createRange();
sel.pasteHTML(arr);
}
}