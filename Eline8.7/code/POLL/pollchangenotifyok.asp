<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<10 or sjjh_name<>application("sjjh_user") then
	Response.Write "<b>[����ʧ��]</b><p>��û�и���վ�������Ȩ�ޣ�"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
notify=RTrim(Request.Form("notify"))
notify=Replace(notify,chr(13)&chr(10),"<br>")
sql="update ͶƱ set �ٸ�����='" & notify & "'"
set rs=conn.Execute(sql)
set rs=nothing
conn.close
set conn=nothing

%><html>
<head>
<title>����վ�������wWw.51eline.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
<!--
.p9 {line-height: 150%; font-size: 9pt;}
.p12 {line-height: 150%; font-size: 12pt;}
body {line-height: 150%;font-size : 12pt;}
td {line-height: 150%;font-size : 12pt;}
A  {text-decoration: none;}
A:Hover  {text-decoration : none;}
a:visited {  color: #0000FF}
-->
</style>
</head>
<body bgcolor="#FFFFEC" topmargin="0" leftmargin="0">
<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#C0C0C0" height="1">
<tr>
    <td width="100%" height="1"><font size="2">����λ��&gt;&gt;<font color="#0000FF"><a href="poll.asp">����ͶƱ</a></font>&gt;&gt;<a href="pollchangenotify.asp">���Ĺٸ����Ź���</a></font></td>
</tr>
</table>
<h1 align="center"><font size="3" color="#0000FF">������վ�����桿</font></h1>
<hr noshade size="1" color=009900>
<p><b><font color="#FF0000">[�������]</font></b></p>
<p>�Ѿ���վ���������Ϊ��<a href="javascript:history.go(-2)">�����ء�</a></p>
<table width="96%" border="1" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<tr>
<td><%=notify%></td>
</tr>
</table>
<hr noshade size="1" color=009900>
</body>
</html>
