<%@ LANGUAGE=VBScript.Encode%>
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Session("yx8_mhjh_userchatroomsn")
allonline=split(Application("yx8_mhjh_allonline"),";")
allonlineubd=ubound(allonline)
%>
<!--#INCLUDE FILE="midsound.asp" -->
<html>
<head>
<title>好歌赠好友</title>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link rel="stylesheet" href="css.css">
<script language="JavaScript">
function gozj(){url="songsend.asp?zj="+document.forms[0].songzj.options[document.forms[0].songzj.selectedIndex].value;window.location.href=url;}
function check(){document.forms[0].send.disabled=1;return true;}
</script>
</head>
<body oncontextmenu=self.event.returnValue=false background="bg1.gif" bgproperties="fixed" text="#FFFFFF">
<div align="center"><font color="#000000" style="font-size:12pt">好歌赠好友</font></div>
<hr size=1 color=#FFFFFF><br>
<table border="0" width="100%">
<form method="post" action="s.asp" onsubmit='return(check());'>
<tr>
<td>
<p><font color="#000000">选择曲目：</font><br>
<select name="song" style="font-size:9pt">
<%For t=1 to 110%> 
<option value="<%=t%>"><%=mid(t)%></option> 
<%next%> 
erase mid 
</select> 
</p> 
<p><font color="#000000">选择好友：</font><br>
<select name="sendto" style="font-size:9pt"><%for i=1 to allonlineubd-1
if allonline(i)<>Application("yx8_mhjh_admin") then
Response.Write "<option value='"&allonline(i)&"'>"&allonline(i)&"</option>"
end if
next
%></select>
</p>
<p> <font color="#000000">祝福话语：</font><br>
<input type="text" name="infofo" size="16" style="font-size:9pt" maxlength="100" value="谨以《%%》表达我的心意……"></p>
<p>注：%%代表歌名
</p>
<table border="0">
<tr>
<td>
<input type="submit" name="send" value="发送">
<input type="button" name="abort" value="放弃" onClick="javascript:history.go(-1)">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
</body>
</html>