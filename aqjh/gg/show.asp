<%
Rem 过滤HTML代码
function DvbbsHtmlAn(fString)
if not isnull(fString) then
	fString = replace(fString, ">", " & gt;")
	fString = replace(fString, "<", " & lt;")

	fString = Replace(fString, CHR(32), " ")
	fString = Replace(fString, CHR(9), " & nbsp;")
	fString = Replace(fString, CHR(34), " & quot;")
	fString = Replace(fString, CHR(39), " & #39;")
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</font></P><P style=""font-size:" & FontSize & "pt;line-height:" & FontHeight & "pt"">")
	fString = Replace(fString, CHR(10), "<BR> ")
	DvbbsHtmlAn = fString
end if
end function
Function FilterJS(v)
if not isnull(v) then
	dim t
    t = Replace(v, "javascript:", "")
    t = Replace(t, "jscript:", "")
    t = Replace(t, "js:", "")
    t = Replace(t, "vbs:", "")
    t = Replace(t, "vbscript:", "")
    FilterJS = t
end if
End Function
function HTMLEncode(fString)
    fString = replace(fString, ">", " & gt;")
    fString = replace(fString, "<", " & lt;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
    fString = Replace(fString, CHR(10), "<BR>")
    HTMLEncode = fString
end function
function UBBCode(strContent)
'	if (Not UserSign and strAllowHTML <> 1) or (UserSign and UserHtmlCode<>1) then
'        strContent = DvbbsHtmlAn(strContent)
'	else
'	strContent = HTMLcode(strContent)
'	end if
'	if (Not UserSign and strAllowForumCode<>1) or (UserSign and UserubbCode<>1) then
'		UBBCode=strContent
'		exit function
'	end if
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
	strContent= re.Replace(strContent,"<br><IMG SRC=""" & picurl & "gif.gif"" border=0>此主题相关图片如下：<br><A HREF=""$2"" TARGET=_blank><IMG SRC=""$2"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")
	re.Pattern="(\[UPLOAD=jpg\])(.[^\[]*)(\[\/UPLOAD\])"
	strContent= re.Replace(strContent,"<br><IMG SRC=""" & picurl & "jpg.gif"" border=0>此主题相关图片如下：<br><A HREF=""$2"" TARGET=_blank><IMG SRC=""$2"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")
	re.Pattern="(\[UPLOAD=bmp\])(.[^\[]*)(\[\/UPLOAD\])"
	strContent= re.Replace(strContent,"<br><IMG SRC=""" & picurl & "bmp.gif"" border=0>此主题相关图片如下：<br><A HREF=""$2"" TARGET=_blank><IMG SRC=""$2"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")

	re.Pattern="(\[UPLOAD=(.[^\[]*)\])(.[^\[]*)(\[\/UPLOAD\])"
	strContent= re.Replace(strContent,"<br><IMG SRC=""" & picurl & "$2.gif"" border=0> <a href=""$3"">点击浏览该文件</a>")

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
	re.Pattern="(\[" & ImgName & "(.[^\[]*)\])"
	strContent=re.Replace(strContent,"<img src=" & picurl&ImgName & "$2.gif border=0 align=middle>")
	end if

	re.Pattern="(\[HTML\])(.[^\[]*)(\[\/HTML\])"
	strContent=re.Replace(strContent,"<table width='100%' border='0' cellspacing='0' cellpadding='6' bgcolor='" & abgcolor & "'><td><b>以下内容为程序代码:</b><br>$2</td></table>")
	re.Pattern="(\[code\])(.[^\[]*)(\[\/code\])"
	strContent=re.Replace(strContent,"<table width='100%' border='0' cellspacing='0' cellpadding='6' bgcolor='" & abgcolor & "'><td><b>以下内容为程序代码:</b><br>$2</td></table>")

	re.Pattern="(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])"
	strContent=re.Replace(strContent,"<font color=$2 style=""font-size:" & FontSize & "pt;line-height:" & FontHeight & "pt"">$3</font>")
	re.Pattern="(\[face=(.[^\[]*)\])(.[^\[]*)(\[\/face\])"
	strContent=re.Replace(strContent,"<font face=$2 style=""font-size:" & FontSize & "pt;line-height:" & FontHeight & "pt"">$3</font>")
	re.Pattern="(\[align=(.[^\[]*)\])(.*)(\[\/align\])"
	strContent=re.Replace(strContent,"<div align=$2>$3</div>")

	re.Pattern="(\[QUOTE\])(.*)(\[\/QUOTE\])"
	strContent=re.Replace(strContent,"<table cellpadding=0 cellspacing=0 border=0 WIDTH=94% bgcolor=#000000 align=center><tr><td><table width=100% cellpadding=5 cellspacing=1 border=0><TR><TD BGCOLOR='" & abgcolor & "'>$2</table></table><br>")
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
	strContent=re.Replace(strContent,"<font size=1 style=""line-height:" & FontHeight & "pt"">$2</font>")
	re.Pattern="(\[size=2\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=2 style=""line-height:" & FontHeight & "pt"">$2</font>")
	re.Pattern="(\[size=3\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=3 style=""line-height:" & FontHeight & "pt"">$2</font>")
	re.Pattern="(\[size=4\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=4 style=""line-height:" & FontHeight & "pt"">$2</font>")
	re.Pattern="(\[size=5\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=5 style=""line-height:" & FontHeight & "pt"">$2</font>")
	re.Pattern="(\[size=10\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=10 style=""line-height:" & FontHeight & "pt"">$2</font>")
	re.Pattern="(\[size=20\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=20 style=""line-height:" & FontHeight & "pt"">$2</font>")
	re.Pattern="(\[size=30\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=30 style=""line-height:" & FontHeight & "pt"">$2</font>")
	re.Pattern="(\[size=40\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=40 style=""line-height:" & FontHeight & "pt"">$2</font>")
	re.Pattern="(\[size=50\])(.[^\[]*)(\[\/size\])"
	strContent=re.Replace(strContent,"<font size=50 style=""line-height:" & FontHeight & "pt"">$2</font>")
	re.Pattern="(\[center\])(.[^\[]*)(\[\/center\])"
	strContent=re.Replace(strContent,"<center>$2</center>")
	set re=Nothing
	UBBCode=strContent
end function
id=request.querystring("id")
if id="" then
response.write "<br><br><div align=center><font color=#ff0000>哈哈，参数错误！</font></div>"
response.end
end if%>
<html><head><title>江湖公告</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.Open "Select * From gg where id=" & id, conn,1,3
RS("c").value = RS("c").value + 1
RS.Update%>
<body background="images/bg.gif"><CENTER><TABLE cellSpacing=0 cellPadding=0 width=80% border=0><TBODY> 
<TR style="FONT-SIZE: 9pt"><TD colspan="2" bgColor=#D7EBFF></TD><TD width=8 height=16></TD></TR>
<TR><TD width=16 bgColor=#D7EBFF></TD><TD style="FONT-SIZE: 9pt; LINE-HEIGHT: 160%" bgColor=#D7EBFF><BLOCKQUOTE>
<P align=center class="3D"><B><%=rs("a")%></B></P><BR>
            <div align="center">类别：<%=rs("d")%> 发布人：(<%=rs("f")%>)　[<%=rs("e")%>]　浏览次数：<%=rs("c")%></div>
<HR color=red SIZE=1><P style="word-break:break-all"><%=ubbcode(rs("b"))%></p>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
<HR color=red SIZE=1><P align=center><input type=button value="关闭窗口" onclick="window.close()" class="input">
<br><br>快乐江湖</P></BLOCKQUOTE>
<P align=center><br></P></TD><TD align=middle width=8 bgColor=#808080></TD></TR>
<TR style="FONT-SIZE: 3pt" align=middle><TD width=16 height=8></TD><TD bgColor=#808080 height=8></TD>
<TD width=8 bgColor=#808080 height=8></TD></TR></TBODY></TABLE></CENTER></body></html>
