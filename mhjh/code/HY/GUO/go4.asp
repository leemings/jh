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
<title>������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></head>
<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="599">
<div align="left"><font color="#FFFFFF"><b>���ߵ��������ؿ�ǰ��ͻȻ����һ���������ι�״С��ͷ~</b></font></div>
</td>
<td width="147" rowspan="2" valign="top"><font color="#FFFFFF"><img src="face.gif" width="142" height="241"></font></td>
</tr>
<tr>
<td width="599" valign="top">
<p style="line-height: 200%; margin: 20"> <font color="#FFFFFF">��������С�ӣ�ˮ׼������Ѿ����˵����ء��ҽд����������꿪ʼ��ܣ��������Ѿ�60�����ˡ��������أ��Ͷ������ȭͷ��������ͷ����ǡ�Ҫ����Ĺ�����߹��ң��Ҿͷ����ȥ~��</font><p><font color="#FFFFFF"><%=username%>:<a href="go5.asp">����ʲô�����书�������µ�һ����</a>
</font>
</p>
<p>
</p>
<p>��</p>
</td>
</tr>
</table>
</body>
</html>



