<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
session("advtime")=now()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
my=aqjh_name
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "�㲻�ǽ������ˣ����ܽ��뽭�����"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
hy=rs("��Ա�ȼ�")
pdhy=rs("��Ա")
if hy<1 and pdhy=False then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('��ѽ��["&aqjh_name &"]�㲻�ǻ�Ա�����������־�ѧ��');window.close();</script>"
 response.end
end if
if rs("����")<10000000 then
		Response.Write "<script language=javascript>alert('��Ǹ�������������');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 then
ss=5-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�ո�����������ȣ�"&ss&"�ֺ����������ɣ�');location.href = 'javascript:history.back()';}</script>"
	Response.End 
end if
tl=rs("����")
nl=rs("����")
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body background="by.gif" bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<center>
<%
if tl<100000 or nl<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "������Ŀǰ������������ֵ��������ȥ������������������ɣ�"
location.href = "javascript:history.back()"
</script>
<%
else
conn.execute "update �û� set ����=����-100000,����=����-50000,����=����-10000000,�书=�书+50000,����ʱ��=now where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
message="" & my & "�������ѧ��һ�죬�����书50000����������100000������50000������1000��"
end if
%>
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr><td height="37">
<font color="#000000"><strong>�������:</strong></font>
<tr>
<td height="182" valign="top"><%=message%><Br><Br><center>
</td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="�� ��" onclick="location.href='index.htm'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>