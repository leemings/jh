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
<title>���Ĺ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></head>
<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="617">
<div align="left"><b><font color="#FFFFFF">���ߵ����Ĺ���ǰ��һ���������ʹ�������ת�˳�������</font></b></div>
</td>
<td width="136" rowspan="2" valign="top"><font color="#FFFFFF"><img src="face2.GIF" width="136" height="283"></font></td>
</tr>
<tr>
<td width="617" valign="top">
<p style="line-height: 200%; margin: 20"> <font color="#FFFFFF">˾ͽ��⣺ɵС�ӣ��봴����˾ͽ�����һ�ؿ�û��ô���ס����������������Ķ�һ�ѣ�Ҫ�������ˣ����Ķ����ľͻ��Ķ�ȥ��Ҫ����Ӯ�ˣ��Ҿͷ����ȥ����ô����Ҫ��Ҫ���ԣ�</font><p style="line-height: 200%; margin: 20"><font color="#FFFFFF"><%=username%>:<a href="go7.asp">���ӿ������µ�һ��������ʲô��</a><br>
</font>
</p>

</td>
</tr>
</table>
</body>
</html>

