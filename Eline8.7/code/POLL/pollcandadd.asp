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
	Response.Write "<b>[����ʧ��]</b><p>��û����Ӻ�ѡ�˵�Ȩ�ޣ�"
	Response.End
end if
addname=Server.HTMLEncode(Trim(Request.Form("addname")))
'addname=Trim(Request.Form("addname"))
if addname="" then
	Response.Write "<b>[����ʧ��]</b><p>��ѡ������Ϊ�գ�������ӣ�<a href=javascript:history.go(-1)>�����ء�</a>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,���� from �û� where ����='" & addname & "'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[����ʧ��]</b><p>�������޴��ˣ�������ӣ�<a href=javascript:history.go(-1)>�����ء�</a>"
	Response.End
end if
id=rs("ID")
rs.close
rs.open "Select ��ѡ�� from ��ѡ�� where ��ѡ��="&id,conn,3,3
if not(rs.eof or rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<b>[����ʧ��]</b><p>�����Ѿ��Ǻ�ѡ���ˣ������ظ���ӣ�<a href=javascript:history.go(-1)>�����ء�</a>"
	Response.End
end if
rs.close
conn.execute "insert into ��ѡ��(��ѡ��,����,��Ʊ��) values ('" & id & "','" & addname & "',0)"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>��Ӻ�ѡ��</title>
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
<table border="1" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#C0C0C0" height="1" cellpadding="0">
<tr>
    <td width="100%" height="1"><font size="2">����λ��&gt;&gt;<font color="#0000FF"><a href="poll.asp">����ͶƱ</a></font><font color="#000000">&gt;&gt;<a href="pollcand.asp">��ѡ�˹���</a>&gt;&gt;</font><font color="#0000FF">��Ӻ�ѡ��</font></font></td>
</tr>
</table>
<h1 align="center"><font color="#0000FF" size="3">����Ӻ�ѡ�ˡ�</font></h1>
<hr noshade size="1" color=009900>
<p><b><font color="#FF0000">[�������]</font></b></p>
<blockquote>
<blockquote>
<p><font size="2">�Ѿ���Ӻ�ѡ�ˣ�<font color="#FF0000"><%=addname%></font> ��<a href="javascript:history.go(-1)">�����ء�</a>
</font>
</p>
</blockquote>
</blockquote>
<hr noshade size="1" color=009900>
</body>
</html>