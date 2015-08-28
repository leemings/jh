<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if Session("usepro")<>"3" then
Session("usepro")=""
response.redirect "index.asp"
end if
%>
<html>
<head>
<title>第四关</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></head>
<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="617">
<div align="left"><b><font color="#FFFFFF">刚走到第四关门前，一个独眼龙就从树后面转了出来……</font></b></div>
</td>
<td width="136" rowspan="2" valign="top"><font color="#FFFFFF"><img src="face2.GIF" width="136" height="283"></font></td>
</tr>
<tr>
<td width="617" valign="top">
<p style="line-height: 200%; margin: 20"> <font color="#FFFFFF">司徒输光：傻小子，想创过我司徒输光这一关可没那么容易。来来来，咱们来赌赌一把，要是你输了，打哪儿来的就回哪儿去；要是你赢了，我就放你过去。怎么样，要不要试试？</font><p style="line-height: 200%; margin: 20"><font color="#FFFFFF"><%=username%>:<a href="go7.asp">老子可是天下第一赌神，怕你什么。</a><br>
</font>
</p>

</td>
</tr>
</table>
</body>
</html>

