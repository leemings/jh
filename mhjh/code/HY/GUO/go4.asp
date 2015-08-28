<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if Session("usepro")<>"2" then
Session("usepro")=""
response.redirect "index.asp"
end if
%>
<html>
<head>
<title>第三关</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></head>
<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="599">
<div align="left"><font color="#FFFFFF"><b>你走到第三个关口前，突然跳出一个打扮的奇形怪状小老头~</b></font></div>
</td>
<td width="147" rowspan="2" valign="top"><font color="#FFFFFF"><img src="face.gif" width="142" height="241"></font></td>
</tr>
<tr>
<td width="599" valign="top">
<p style="line-height: 200%; margin: 20"> <font color="#FFFFFF">打不死：臭小子，水准不错嘛。已经到了第三关。我叫打不死，从三岁开始打架，到今天已经60多年了。想过我这关，就动动你的拳头，让我老头活活筋骨。要是你的攻击里高过我，我就放你过去~！</font><p><font color="#FFFFFF"><%=username%>:<a href="go5.asp">怕你什么，我武功可是天下第一！！</a>
</font>
</p>
<p>
</p>
<p>　</p>
</td>
</tr>
</table>
</body>
</html>



