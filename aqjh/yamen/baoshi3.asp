<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
myid=Request.form("id")
name=request.form("name")
pass=trim(request.form("pass"))
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��鿴��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
if name="" or pass="" then
		Response.Write "<script Language=Javascript>alert('��ʾ���ǲ�����һ�������η�ѽ���������ͽ��ſ��������');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
inroom=i
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where id="& myid &" and ����='" & name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����û�и������������˰���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if trim(pass)<>rs("����") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������벻�ԣ���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("����")<30000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������û����ô��Ǯѽ����ô�죿');}</script>"
	Response.End
end if
if rs("״̬")<>"��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����û�и������������û��������ˣ���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
conn.execute("update �û� set ״̬='����' where ����='"&name&"'")
conn.execute "update �û� set ����=����-30000000 where ����='"&name&"'"
Response.Write "<script Language=Javascript>alert('OK,���ó���3000�������������ұ�����,�Ժ��ٸɵ㻵�°�!');window.close();</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>