<%@ LANGUAGE=VBScript.Encode%>
<!--#INCLUDE FILE="midsound.asp" -->
<%
if Request.form("play")="播放" then
song=Request.form("song")
loopok=Request.form("loopok")
linkurl="mid/" & song &".mid"
 response.write "<script language=javascript>parent.topfrm.location.replace('welcome.asp?search="&song&"&do="&loopok&"');</script>"
end if
if Request.form("st")="停止" then
song=""
loopok=""
linkurl=""
 response.write "<script language=javascript>parent.topfrm.location.replace('welcome.asp?search="&song&"&do="&loopok&"');</script>"
end if
%>
<html>
<head>
<title>点歌</title>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<script language="javaScript">
function send(){sendurl="songsend.asp?songname=" + document.forms[0].song.options[document.forms[0].song.selectedIndex].value;window.location.href=sendurl;}
</script>
<link rel="stylesheet" href="css.css">
</head>
<body oncontextmenu=self.event.returnValue=false  background="bg1.gif" text="#FFFFFF">
<div align=center><font color="#000000" style="font-size:12pt">点歌系统</font></div>
<hr size=1 color=#FFFFFF>
<table border="0" width="100%">
<form method="post" action="song.asp" name="" >
<tr>
<td>
<div>
<p><a href="javascript:history.go(0)" title="显示最新的曲目列表">刷新:<font color="#FFFFFF"></font></a></p>
</div>
<p> 选择曲目：<br>
<select name="song" style="font-size:9pt">
<%For t=1 to 110%> 
<option value="<%=t%>"><%=mid(t)%></option> 
<%next%> 

 
</select> 
</p> 
<p>播放方式：<br> 
<input type="radio" name="loopok" value="1" checked> 
只听一遍<br> 
<input type="radio" name="loopok" value="infinite"> 
百听不厌</p> 
<table border="0" cellpadding="4"> 
<tr> 
<td> 
<input type="submit" name="play"  value="播放"> 
</td> 
<td align="right"> 
<input type="submit" name="st" value="停止"> 
</td> 
</tr> 
<tr> 
<td colspan="2"> 
<input type="button" value="好歌赠好友" onClick="Javascript:send()" name="button"> 
</td> 
</tr> 
</table> 
</td> 
</tr> 
</form> 
</table> 
</body> 
</html>

