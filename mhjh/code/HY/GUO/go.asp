<!--#include file="../../config.asp"--><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT ��Ա,���� FROM �û� where ����='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "�㲻�ǽ������˻������ӳ�ʱ"
conn.close
response.end
else
hy=rs("��Ա")
if username="" then
%>
<script language=vbscript>
MsgBox "�Բ����㻹û�е�¼"
location.href = "../../exit.asp"
</script>
<%
else
if yx8_mhjh_fellow=false then
%>
<script language=vbscript>
MsgBox "���󣡻�Ա������δ���ţ�"
location.href = "javascript:history.back()"
</script>
<%
else
if yx8_mhjh_fellow=false then
%>
<script language=vbscript>
MsgBox "���󣡻�Ա������δ���ţ�"
location.href = "javascript:history.back()"
</script>
<%
end if
if rs("����")>15000000 then
%>
<script language=vbscript>
MsgBox "�������Ѿ��Ƕ��������ˣ���ذɣ�"
location.href = "javascript:history.back()"
</script>
<% else %>
<html>
<head>
<title>��һ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></head>

<body bgcolor="#000000" oncontextmenu=self.event.returnValue=false >
<table width="562" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="595">
<div align="left"><font color="#FFFFFF"><b>���صĵ�һ���Ƕ�����ܵ�ֿ�ѡ���������ϵ�</b></font></div>
</td>
<td width="162" rowspan="2" valign="top"><img src="jh1.jpg" width="160" height="250"></td>
</tr>
<tr>
<td width="595" valign="top">
<p style="line-height: 200%; margin: 20"> <font color="#FFFFFF"><font color=red>���ϵ���</font>��С�ӣ������Ǳ��صĵ�һ�أ������˼������;Ͳ����㶯���ˣ����������壬��Ӯ���Ҿͷ����ȥ������Ҫ���ˣ��Ҿ�ֱ�������ɽ�ڣ�</font><br><br>
<font color=red><%=username%></font>:<a href="go1.asp">�ðɣ���һ������Ӯ��ġ�</a><br>


</p></td>
</tr>
</table>
</body>
</html>
<% 
end if
end if
end if
end if
rs.Close              
set rs=nothing              
conn.Close              
set conn=nothing 
 %>