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
<title>���һ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></head>
<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="630">
<div align="left"><b><font color="#FFFFFF">���һ���ˣ����ص���ǰ����һ�������ˡ���</font></b></div>
</td>
<td width="128" rowspan="2" valign="top"><img src="face3.GIF" width="128" height="290"></td>
</tr>
<tr>
<td width="630" valign="top">
<p style="line-height: 200%; margin: 20"> <font color="#FFFFFF">�����ˣ���������˭~����ֻҪ֪������������һ�����־Ϳ����ˡ�ǰ���4�أ��㾭���������������������������Ŀ��飬��������������������������ܰ��һ��ܣ������������ƿ����صĴ��ţ���������ɣ��Ҳ�������ģ�</font>      <p style="line-height: 200%; margin: 20"><font color="#FFFFFF"><%=session("yx8_mhjh_username")%>:<a href="go9.asp">��~׼��������Ұɣ�����</a><br>
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


