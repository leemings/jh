<%@ LANGUAGE="VBSCRIPT" %>
<%
dim UserName,Password
UserName=request.form("UserName")
UserName=replace(UserName,"'","''")
Password=request.form("Password")
Password=replace(Password,"'","''")
if len(UserName)=0 or len(Password)=0 then
response.write"<Script Language=Javascript>window.alert('��������д��¼��Ϣ��\n\n���ȷ�������������롣');history.back()</Script>"
response.end
end if
%>
<!--#include file="conn.asp"-->
<%
dim rs,sql
sql="Select * From Admin Where UserName='"&UserName&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
conn.close
set rs=nothing
set conn=nothing
response.write"<Script Language=Javascript>window.alert('����Ա�ʺ���������\n\n���ȷ�������������롣');history.back()</Script>"
response.end
elseif rs("Password")<>Password or len(rs("Password"))<>len(Password) or rs("UserName")<>UserName or len(rs("UserName"))<>len(Username) then
conn.close
set rs=nothing
set conn=nothing
response.write"<Script Language=Javascript>window.alert('����Ա������������\n\n���ȷ�������������롣');history.back()</Script>"
response.end
end if
session("Level")=rs("Level")
session("UserName")=UserName
set rs=nothing
set conn=nothing
response.write"<Script Language=Javascript>window.alert('��ϲ������¼�ɹ���\n\n���ȷ���������ҳ�档');location.href='main.asp';</Script>"
response.end
%>