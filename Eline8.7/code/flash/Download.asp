<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ��������Flash���Ƚ��뽭�������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT ����,���� FROM �û� WHERE ����='"& session("sjjh_name") &"'",conn,1,3
if rs("����")<100000 or rs("����")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "����û100000���������Ҳ���1������������Flash��"
	response.end
end if
rs("����")=rs("����")-100000
rs("����")=rs("����")-1
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="update flash set download=download+1 where ID="&request("ID")
rs.open sql,conn,1,3 
sql="select * from flash where id="&request("id")         
rs.open sql,conn,1,1
url=""&rs("url")&""
response.redirect ""&url&""
response.end
rs.close                               
set rs=nothing                               
conn.close                               
set conn=nothing
%>