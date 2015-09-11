<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=request("id")
wj=request("wj")
if id="" or (not isnumeric(id)) then 
	Response.Write "<script language=JavaScript>{alert('提示：你在搞什么！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
set conn1=server.createobject("adodb.connection")
conn1.open Application("aqjh_usermdb")
set rs1=server.CreateObject ("adodb.recordset")
rs1.Open "Select * From 日记 where id="& id, conn1, 1,1
if rs1.eof and rs1.bof then
	call diaryclose("此日记不存在，请刷新浏览器！",wj,aqjh_name)
end if
name=rs1("用户名")
mm=rs1("保密")
select case mm
case "1"
case "2"
	if name<>aqjh_name then
		call diaryclose("私人保密，禁止观看！",wj,name)
	end if
case "3"
	pass=trim(request("pass"))
	if pass="" and name<>aqjh_name then
		Response.Write "<script Language=Javascript>var pass=prompt("& chr(34) &"请输入访问密码："& chr(34) &"," &chr(34)&chr(34) &");location.href = "&chr(34) &"show.asp?id="& id &"&pass="& chr(34) &"+pass;</script>"	
	else
		if pass<>rs1("保密条件") and name<>aqjh_name then
			call diaryclose("输入的密码不正确!",wj,name)
		end if
	end if
case "4"
	if aqjh_name<>rs1("保密条件") and name<>aqjh_name then
		call diaryclose("只有["& rs1("保密条件") &"]才可以查看!",wj,name)
	end if
case "5"
	if aqjh_jhdj<clng(rs1("保密条件")) and name<>aqjh_name then
		call diaryclose("限制战斗等级["& rs1("保密条件") &"]级可查看!",wj,name)
	end if
case "6"
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	rs.open "select 门派 from 日记用户 where 姓名='"& aqjh_name &"'",conn,3,3
	mp=rs("门派")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	if mp<>rs1("保密条件") and name<>aqjh_name then
		call diaryclose("限制门派：["& mp &"]才可观看!",wj,name)
	end if
case else
		call diaryclose("日记数据有错，不能查看!",wj,name)
end select
function badstr(str)
	badword="射精,奸,屎,妈,娘,尻,操,王八,逼,贱,狗,婊,叉你,插你,干你,鸡巴,睾丸,蛋,包皮,龟头,屄,赑,妣,肏,奶,尻,屌,作爱,做爱,床,抱抱,鸡八,处女,打炮,十八摸,爸,我儿,·,主席,泽民,法伦,洪志,大法,公安,政府,反动,法院,升天,周天,功"
	bad=split(badword,",")
	for i=0 to ubound(bad)-1
		if InStr(LCase(str),bad(i))<>0 then 
			str=replace(str,bad(i),"*")
		end if
	next
	badstr=str
end function
sub diaryclose(mess,wj,myname)
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	if wj="list.asp" then
		Response.Write "<script Language=Javascript>alert('提示："& mess &"');window.close();</script>"
	else
		Response.Write "<script Language=Javascript>alert('提示："& mess &"');window.close();</script>"
	end if
	response.end
end sub
aa1=rs1("用户名")
aa2=badstr(rs1("日记标题"))
aa3=badstr(UBBCode(rs1("日记内容")))
aa4=rs1("天气")
aa5=rs1("心情")
aa8=rs1("书写日期")
aa9=rs1("鲜花")
aa10=rs1("鸡蛋")
aa11=rs1("留言数")
'aa12=rs1("留言")
aa13=rs1("lb")
rs1.close
rs1.Open "Select * From 日记用户 where 姓名='"& aa1 &"'", conn1, 1,1
if rs1.eof and rs1.bof then
	call diaryclose("此用户资料有问题，无法查看！","list.asp",name)
end if
bb0=badstr(rs1("日记本名"))
bb4=rs1("鲜花")
bb5=rs1("鸡蛋")
bb6=rs1("日记数")
bb7=rs1("留言数")
bb8=rs1("建立日期")
bb9=rs1("文采")
dim mylb(3)
mylb(1)=rs1("lb1")
mylb(2)=rs1("lb2")
mylb(3)=rs1("lb3")

rs1.close
set rs1=nothing
conn1.close
set conn1=nothing
function DvbbsHtmlAn(fString)
if not isnull(fString) then
	fString = replace(fString, ">", "&gt;")
	fString = replace(fString, "<", "&lt;")
	fString = Replace(fString, CHR(32), " ")
	fString = Replace(fString, CHR(9), "&nbsp;")
	fString = Replace(fString, CHR(34), "&quot;")
	fString = Replace(fString, CHR(39), "&#39;")
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</font></P><P style=""font-size:"&FontSize&"pt;line-height:"&FontHeight&"pt"">")
	fString = Replace(fString, CHR(10), "<BR> ")
	DvbbsHtmlAn = fString
end if
end function
function UBBCode(strContent)
	dim re
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern="(\[IMG\])((javascript:|jscript:|js:|vbs:|vbscript:).[^\[]*)(\[\/IMG\])"
	strContent=re.Replace(strContent,"$2")
	re.Pattern="(\[UPLOAD=((javascript:|jscript:|js:|vbs:|vbscript:).[^\[]*)\])(.[^\[]*)(\[\/UPLOAD\])"
	strContent= re.Replace(strContent,"$3")
	re.Pattern="(\[UPLOAD=(.[^\[]*)\])((javascript:|jscript:|js:|vbs:|vbscript:).[^\[]*)(\[\/UPLOAD\])"
	strContent= re.Replace(strContent,"$3")
	if (Not UserSign and strIMGInPosts = 1) or (UserSign and UserImgCode=1) then
	re.Pattern="(\[IMG\])(http://.[^\[]*)(\[\/IMG\])"
	strContent=re.Replace(strContent,"<IMG SRC=""$2"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333"">")
	end if
	if Not UserSign or (UserSign and UserImgCode=1) then
	re.Pattern="\[DIR=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/DIR]"
	strContent=re.Replace(strContent,"<object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=$1 height=$2><param name=src value=$3><embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed></object>")
	re.Pattern="\[QT=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/QT]"
	strContent=re.Replace(strContent,"<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>")
	re.Pattern="\[MP=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/MP]"
	strContent=re.Replace(strContent,"<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3  width=$1 height=$2></embed></object>")
	re.Pattern="\[RM=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/RM]"
	strContent=re.Replace(strContent,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")
	end if
	if (Not UserSign and strflash= 1) or (UserSign and UserImgCode=1) then
	re.Pattern="(\[FLASH\])(.[^\[]*)(\[\/FLASH\])"
	strContent= re.Replace(strContent,"<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")
	re.Pattern="(\[FLASH=*([0-9]*),*([0-9]*)\])(.[^\[]*)(\[\/FLASH\])"
	strContent= re.Replace(strContent,"<a href=""$4"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画!> [全屏欣赏]</a><br><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE=""$4""><PARAM NAME=quality VALUE=high><embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")
	end if
	re.Pattern="(\[UPLOAD=gif\])(.[^\[]*)(\[\/UPLOAD\])"
	strContent= re.Replace(strContent,"<br><IMG SRC="""&picurl&"gif.gif"" border=0>此主题相关图片如下：<br><A HREF=""$2"" TARGET=_blank><IMG SRC=""$2"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")
	re.Pattern="(\[UPLOAD=jpg\])(.[^\[]*)(\[\/UPLOAD\])"
	strContent= re.Replace(strContent,"<br><IMG SRC="""&picurl&"jpg.gif"" border=0>此主题相关图片如下：<br><A HREF=""$2"" TARGET=_blank><IMG SRC=""$2"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")
	re.Pattern="(\[UPLOAD=bmp\])(.[^\[]*)(\[\/UPLOAD\])"
	strContent= re.Replace(strContent,"<br><IMG SRC="""&picurl&"bmp.gif"" border=0>此主题相关图片如下：<br><A HREF=""$2"" TARGET=_blank><IMG SRC=""$2"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")
	re.Pattern="(\[UPLOAD=(.[^\[]*)\])(.[^\[]*)(\[\/UPLOAD\])"
	strContent= re.Replace(strContent,"<br><IMG SRC="""&picurl&"$2.gif"" border=0> <a href=""$3"">点击浏览该文件</a>")
	re.Pattern="(\[URL\])(.[^\[]*)(\[\/URL\])"
	strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")
	re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
	strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")
	re.Pattern="(\[EMAIL\])(.[^\[]*)(\[\/EMAIL\])"
	strContent= re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$2"">$2</A>")
	re.Pattern="(\[EMAIL=(.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
	strContent= re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$2"" TARGET=_blank>$3</A>")
	re.Pattern = "^(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "[^>=""](http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "^(ftp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "(ftp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "[^>=""](ftp://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "^(rtsp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "(rtsp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "[^>=""](rtsp://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "^(mms://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "(mms://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	re.Pattern = "[^>=""](mms://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif><a target=_blank href=$1>$1</a>")
	if strIcons = "1" then
	re.Pattern="(\["&ImgName&"(.[^\[]*)\])"
	strContent=re.Replace(strContent,"<img src="&picurl&ImgName&"$2.gif border=0 align=middle>")
	end if
	re.Pattern="(\[HTML\])(.[^\[]*)(\[\/HTML\])"
	strContent=re.Replace(strContent,"<table width='100%' border='0' cellspacing='0' cellpadding='6' bgcolor='"&abgcolor&"'><td><b>以下内容为程序代码:</b><br>$2</td></table>")
	re.Pattern="(\[code\])(.[^\[]*)(\[\/code\])"
	strContent=re.Replace(strContent,"<table width='100%' border='0' cellspacing='0' cellpadding='6' bgcolor='"&abgcolor&"'><td><b>以下内容为程序代码:</b><br>$2</td></table>")
	re.Pattern="(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])"
	strContent=re.Replace(strContent,"<font color=$2 style=""font-size:"&FontSize&"pt;line-height:"&FontHeight&"pt"">$3</font>")
	re.Pattern="(\[face=(.[^\[]*)\])(.[^\[]*)(\[\/face\])"
	strContent=re.Replace(strContent,"<font face=$2 style=""font-size:"&FontSize&"pt;line-height:"&FontHeight&"pt"">$3</font>")
	re.Pattern="(\[align=(.[^\[]*)\])(.*)(\[\/align\])"
	strContent=re.Replace(strContent,"<div align=$2>$3</div>")
	re.Pattern="(\[QUOTE\])(.*)(\[\/QUOTE\])"
	strContent=re.Replace(strContent,"<table cellpadding=0 cellspacing=0 border=0 WIDTH=94% bgcolor=#000000 align=center><tr><td><table width=100% cellpadding=5 cellspacing=1 border=0><TR><TD BGCOLOR='"&abgcolor&"'>$2</table></table><br>")
	re.Pattern="(\[fly\])(.*)(\[\/fly\])"
	strContent=re.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$2</marquee>")
	re.Pattern="(\[move\])(.*)(\[\/move\])"
	strContent=re.Replace(strContent,"<MARQUEE scrollamount=3>$2</marquee>")	
	re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
	strContent=re.Replace(strContent,"<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
	re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
	strContent=re.Replace(strContent,"<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")
	re.Pattern="(\[i\])(.[^\[]*)(\[\/i\])"
	strContent=re.Replace(strContent,"<i>$2</i>")
	re.Pattern="(\[u\])(.[^\[]*)(\[\/u\])"
	strContent=re.Replace(strContent,"<u>$2</u>")
	re.Pattern="(\[b\])(.[^\[]*)(\[\/b\])"
	strContent=re.Replace(strContent,"<b>$2</b>")
	re.Pattern="(\[size=1\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=1 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=2\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=2 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=3\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=3 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=4\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=4 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=5\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=5 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=10\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=10 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=20\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=20 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=30\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=30 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=40\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=40 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=50\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=50 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[center\])(.[^\[]*)(\[\/center\])"
	strContent=re.Replace(strContent,"<center>$2</center>")
	set re=Nothing
	UBBCode=strContent
end function
%>
<html><head><title>查看[<%=aa1%>]的日记，标题：<%=aa2%></title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<SCRIPT language=JavaScript>
function textCounter(field, countfield, maxlimit) {
if (field.value.length > maxlimit)field.value = field.value.substring(0, maxlimit);
else countfield.value = maxlimit - field.value.length;}
function check()
{var titly=document.myform.titly.value;
var myly=document.myform.myly.value;
if(titly=="" || myly=="" ){alert("提示：数据不能为空！");return false;}
if(titly.length>20){alert("提示：标题文字太多！");return false;}
if(myly.length>150){alert("提示：留言内容太多！");return false;}
if(titly.length<1){alert("提示：标题文字太少！");return false;}
if(myly.length<5){alert("提示：留言内容太少！");return false;}}
</script>
</head><LINK href="images/diary.css" rel=stylesheet type=text/css><body bgcolor="#FFFFFF" text="#000000">
<table align=center border=0 cellpadding=0 cellspacing=0 width=574><tr><td><img height=13 src="images/top.gif" width=574></td></tr><tr>
    <td align=right background=images/top1.gif height="228"> 
      <table border=0 cellpadding=0 cellspacing=0 width="94%" height="332">
<tr><td align=right><table border=0 cellpadding=0 cellspacing=0 width="80%"><tr align=right><td><a href="mydiary.asp"><img alt=进入我的日记本 border=0 height=21 src="images/anniu1.gif" width=81></a></td>
<td><a href="list.asp"><img alt=最新日记 border=0 height=21 src="images/anniu2.gif" width=80></a></td>
<td><a href="top.asp?fs=1"><img border=0 height=21 src="images/anniu3.gif" width=91></a></td>
<td><a href="top.asp?fs=2"><img alt=日记排行 border=0 height=21 src="images/anniu4.gif" width=78></a></td></tr>
</table></td><td rowspan=3 valign=bottom width=26><img height=200 src="images/right.gif" width=8></td></tr>
<tr><td><div align="left"><HR color=red SIZE=1><a href="list.asp?fs=1&name=<%=aa1%>" target="_blank"><img src="images/hot.gif" width="11" height="14" border="0" alt="查看作者专辑！"></a>[<%=aa1%>]创建于：<font color=blue><%=year(bb8)%>年<%=month(bb8)%>月<%=day(bb8)%>日</font> 总数：<font color=blue><%=bb6%></font> 鲜花：<font color=blue><%=bb4%></font>
鸡蛋：<font color=blue><%=bb5%></font> 留言：<font color=blue><%=bb7%></font> 文采:<font color=blue><%=bb9%></font></div>
</td></tr><tr>
          <td align=middle valign=top height="166"> 
            <table border=0 cellpadding=0 cellspacing=0 width="100%" bordercolor="#3333CC"><tr>
<td colspan="2"><table border=0 cellpadding=3 cellspacing=0 width="100%"><tr><td align=middle valign=top>
<div align="right">日记本名：<font color=blue><%=bb0%></font> 专辑：『<a href="list.asp?fs=2&name=<%=aa1%>&zz=1" ><font color=blue><%=mylb(1)%></font></a> 』『<a href="list.asp?fs=2&name=<%=aa1%>&zz=2" ><font color=blue><%=mylb(2)%></font></a> 』『<a href="list.asp?fs=2&name=<%=aa1%>&zz=3" ><font color=blue><%=mylb(3)%></font></a>』</div>
</td></tr></table></td></tr><tr bgcolor="#66CCFF"><td align=middle colspan="2" ><P align=left class="3D">标题：<%=aa2%></p>
</td></tr><tr><td align=middle height=13 colspan="2"><div align="center">日期：<%=aa8%> 天气：<img src="images/wether<%=aa4%>.gif" width="19" high="19">
心情：<img src="images/<%=aa5%>.gif" width="19" high="19">  所属专辑：『<%=mylb(aa13)%> 』</div>
</td></tr><tr><td align=middle colspan="2"><p align="left" >内容：<br><%=aa3%> </p><p align="right">By <%=aa1%> 于 <%=aa8%>书&nbsp;&nbsp;<br><br> </p>
</td></tr><tr><td align=right height=20 width="40%"><div align="center"><img src="images/xianhua.gif" width="16" height="15">:<font color=blue><%=aa9%></font> 
                    <img src="images/jidan.gif" width="16" height="15">:<font color=blue><%=aa10%></font> 
                    留言：<font color=blue><%=aa11%></font></div>
</td>
                <td align=right height=20 width="60%">送束<img src="images/xianhua.gif" width="16" height="15"> 
                  送个<img src="images/jidan.gif" width="16" height="15"> 
                  <%if aqjh_name=aa1 or aqjh_grade=10 then%>
                  <a href="diary.asp?action=del&id=<%=id%>">删除日记</a> <a href="edit.asp?id=<%=id%>">编辑日记</a> 
                  <a href="diary.asp?action=delly&id=<%=id%>">删全部留言</a> 
                  <%end if%>
                </td>
</tr></table><hr color=red size=1 width="50%" align="center">
            <table width="100%" border="0" height="32" cellspacing="0" cellpadding="0">
              <form action="diary.asp?action=ly&id=<%=id%>" method=post name=myform onSubmit="return check(this);">
                <tr>
                  <td height="77" width="58%" rowspan="2" valign="top"><font color=blue>留言列表:</font><br>
                    <script src="lylist.asp?id=<%=id%>"></script>
                  </td>
                  <td height="8" valign="top" width="11%">主题：</td>
                  <td height="8" valign="top" width="31%"> 
                    <input  name=titly size=22 maxlength="14">
                  </td>
                </tr>
                <tr>
                  <td height="2" valign="top" width="11%">内容：</td>
                  <td height="2" valign="top" width="31%">
                    <textarea  cols=22 name=myly onKeyDown=textCounter(this.form.myly,this.form.remLen,150); onKeyUp=textCounter(this.form.myly,this.form.remLen,150); rows=3 wrap="VIRTUAL" ></textarea>
                    输入
                    <input maxlength=3 name=remLen readOnly size=3  value="150" style="BORDER-BOTTOM-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-RIGHT-WIDTH: 0px; BORDER-TOP-WIDTH: 0px; COLOR: red" >
                    个字符 
                    <input id=submit name=publish  type=submit value=发表>
                  </td>
                </tr>
              </form>
            </table>
            <div align="center">
            <%if wj="list.asp" then
				Response.Write "<a href=list.asp>返回上一级</a>"
			else
				Response.Write "<a href=main.asp?name="& name &">返回上一级</a>"
			end if%></div>
          </td>
        </tr></table>
</td></tr><tr>
    <td height="2"><img height=43 src="images/down.gif" width=574></td>
  </tr></table></body></html>
