<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
my=sjjh_name%>
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
conn.open Application("sjjh_usermdb")
rs.open "select * from �û� where ����='"&sjjh_name&"'",conn
sql="update �û� set ����=����+100,����=����+10000 where ����='"&sjjh_name&"'"
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