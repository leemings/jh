<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if Session("usepro")<>"4" then
Session("usepro")=""
response.redirect "index.asp"
end if
%>
<html>
<head>
<title>最后一关</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></head>
<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="630">
<div align="left"><b><font color="#FFFFFF">最后一关了，宝藏的门前坐着一个年轻人……</font></b></div>
</td>
<td width="128" rowspan="2" valign="top"><img src="face3.GIF" width="128" height="290"></td>
</tr>
<tr>
<td width="630" valign="top">
<p style="line-height: 200%; margin: 20"> <font color="#FFFFFF">年轻人：别问我是谁~！你只要知道我是你的最后一个对手就可以了。前面的4关，你经过了棋力、智力、攻击、赌术的考验，现在有我来考验你的内力，你能把我击败，才能用内力破开宝藏的大门，放马过来吧，我不会留情的！</font>      <p style="line-height: 200%; margin: 20"><font color="#FFFFFF"><%=session("yx8_mhjh_username")%>:<a href="go9.asp">好~准备好输给我吧！！！</a><br>
</font>
</p>
<p style="line-height: 200%; margin: 20">
</td>
</tr>
<tr valign="top">
<td colspan="2">
</td>
</tr>
</table>
</body>
</html>


