<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
myid=Request.form("id")
name=Request.form("name")
oldpass=Request.form("oldpass")
pass=request.form("pass")
repass=request.form("repass")
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��鿴��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
if not isnumeric(myid) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&myid&"]�������ID��ʹ�����֣�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if name="" then
message="�ʺŲ���Ϊ��"
elseif oldpass="" or pass="" or repass="" then
message="���벻��Ϊ��"
elseif pass<>repass then
message="������ȷ���������һ��"
else
oldpass=md5(oldpass)
pass=md5(pass)
'ip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
'У���û�
rs.open "SELECT ״̬ FROM �û� WHERE id="&myid&" and ����='" & name & "' and ����='" & oldpass& "'",conn
If Rs.Bof OR Rs.Eof Then
message="�Բ���������벻��"
else
	if  rs("״̬")<>"����" then
		message="���״̬�����������Ը����룡"
	else
		conn.Execute "update �û� set ����='" & pass & "' where ����='" & name & "'"
		conn.execute "insert into l(a,b,c,d,e) values (now(),'"& name &"','��','����','�Ļ��û����룡')"
		message="��ϲ���ɹ����޸������룡"
	end if
end if
	conn.close
	set conn=nothing
end if
%>
<head><LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699">
<table width="360" border="1" align="center" cellpadding="10"
cellspacing="13" bgcolor="#000000">
  <tr bgcolor="#FFFFFF">
<td>
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=message%></p>
<p align="center"><a href="#" onclick="window.close()">[ �� �� �� �� ]</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>