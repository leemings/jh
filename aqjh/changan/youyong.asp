<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
my=aqjh_name%>
<!--#include file="data1.asp"-->
<%sql="select * from ���� where ����='" & my & "' or ����='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>
MsgBox "����û�������ء�"
location.href = "fangwu.asp"
</script>
<%else
if day(rs("��Ӿ"))<day(now()) and month(rs("��Ӿ"))<month(now()) and year(rs("��Ӿ"))<year(now()) and isnull(rs("��Ӿ")) then
sql="update ���� set ��Ӿ=now() where ����='"& my &"' or ����='" & my & "'"
set rs=conn.execute(sql)
set rs=nothing	
	conn.close
	set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
sql="update �û� set ����=����+100,����=����+10000 where ����='"&aqjh_name&"'"
set rs=conn.execute(sql)
Response.Redirect "../ok.asp?id=602"
else
%><script language=vbscript>
MsgBox "�������ι�Ӿ�˰���"
location.href = "xiaowu7.asp"
</script>
<%end if
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>