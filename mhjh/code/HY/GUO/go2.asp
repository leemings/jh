<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if Session("usepro")<>"1"  then
Session("usepro")=""
response.redirect "index.asp"
end if
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
%>
<html>
<head>
<title>第二关</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></head>
<body bgcolor="#000000" oncontextmenu=self.event.returnValue=false >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="595">
<div align="left"><b><font color="#FFFFFF">第二关是一位PPMM耶~！！！</font></b></div>
</td>
<td width="165" rowspan="2" valign="top"><font color="#FFFFFF"><img src="jh0.jpg" width="153" height="172"></font></td>
</tr>
<tr>
<td width="595" valign="top">
<p style="line-height: 200%; margin: 20"> <font color="#FFFFFF">PPMM：嗨，小帅哥，我父亲是独孤城的智囊。我来代我的父亲把手第二关，不要怕呀，我这里有十道问题，全部答对了就可以过去，要不要试试呀？</font></p>
</td>
</tr>
<tr valign="top">
<td colspan="2">
<p><font color=red><%=username%>:</font><a href="answer.asp">好吧，我的外号可是叫做“问题天才”，把卷子给我。</a><br>
</p>
</td>
</tr>
</table>
</body>
</html>




