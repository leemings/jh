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
	Response.Write "<b>[����ʧ��]</b><p>��û�и���ͶƱʱ���Ȩ�ޣ�"
	Response.End
end if
begintime=Trim(Request.Form("begintime"))
endtime=Trim(Request.Form("endtime"))
if IsDate(begintime)=false or len(begintime)<>18 then
	Response.write "<b>[����ʧ��]</b><p>��ʼͶƱʱ�䣺<font color=FF0000>" & begintime & "</font> ��ʽ����<a href=javascript:history.go(-1)>�����ء�</a>"
	Response.end
end if
if IsDate(endtime)=false or len(endtime)<>18 then
	Response.write "<b>[����ʧ��]</b><p>��ֹͶƱʱ�䣺<font color=FF0000>" & endtime & "</font> ��ʽ����<a href=javascript:history.go(-1)>�����ء�</a>"
	Response.end
end if
if CDate(endtime)<=CDate(begintime) then
	Response.write "<b>[����ʧ��]</b><p><font color=FF0000>��ֹͶƱʱ��</font> �������� <font color=FF0000>��ʼͶƱʱ��</font>��<a href=javascript:history.go(-1)>�����ء�</a>"
	Response.end
end if
if CDate(endtime)<=now() then
	Response.write "<b>[����ʧ��]</b><p><font color=FF0000>��ֹͶƱʱ��</font> �������� <font color=FF0000>��ǰʱ��</font>��<a href=javascript:history.go(-1)>�����ء�</a>"
	Response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="update ͶƱ set ��ʼʱ��='" & begintime & "',��ֹʱ��='" & endtime & "'"
set rs=conn.execute(sql)
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>����ͶƱ��ֹʱ��</title>
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
<body bgcolor="#FFFFF4" topmargin="0" leftmargin="0">
<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#C0C0C0" height="1">
<tr>
    <td width="100%" height="1"><font size="2">����λ��&gt;&gt;<font color="#0000FF"><a href="poll.asp">����ͶƱ</a></font>&gt;&gt;<a href="pollreset.asp">����ͶƱʱ��</a></font></td>
</tr>
</table>
<h1 align="center"><font size="3" color="#0000FF">������ͶƱ��ֹʱ�䡿</font></h1>
<hr noshade size="1" color=009900>
<blockquote>
<p><b><font size="3" color="#FF0000">[�������]</font></b></p>
<blockquote>
<p><font size="2">�Ѿ���ͶƱʱ�����Ϊ��<a href="javascript:history.go(-2)">�����ء�</a></font></p>
<p><font size="2">��ʼͶƱʱ�䣺<font color=FF0000><%=begintime%></font></font></p>
<p><font size="2">��ֹͶƱʱ�䣺</font><font color=FF0000><%=endtime%></font></p>
</blockquote>
</blockquote>
<hr noshade size="1" color=009900>
</body>
</html>
