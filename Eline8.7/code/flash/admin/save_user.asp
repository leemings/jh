<!--#include file="session.asp"-->
<!--#include file="level.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/char.inc"-->
<%
dim Username,Password,Level,Userid,action
Username=request.form("username")
Password=request.form("Password")
Level=request.form("Level")
Userid=request.form("id")
action=request.form("action")
%>
<%
if len(Username)=0 or Level="" or len(Password)=0 then
response.write"<Script>alert('�û������û����롢�û��ȼ�����Ϊ��');history.back()</Script>"
response.end
end if
if len("Username")>20 or len("Password")>20 then
response.write"<Script>alert('�û������û��������20���ַ�');history.back()</Script>"
response.end
end if
if action="" then
action="add"
end if
%>
<%
if action="add" then
Username=replace(request.form("Username"),"'","''")
Password=replace(request.form("Password"),"'","''")
set rs=server.createobject("adodb.recordset")
sql="select * from Admin"
rs.open sql,conn,1,3
rs.addnew
rs("Username")=Username
rs("Password")=Password
rs("Level")=Level
rs.update
rs.close
set rs=nothing
response.write"<Script>alert('���û���ӳɹ�');location.href='add_user.asp';</Script>"
response.end
elseif action="edit" then
Username=replace(request.form("Username"),"'","''")
Password=replace(request.form("Password"),"'","''")
set rs=server.createobject("adodb.recordset")
sql="select * from Admin where id="&Userid
rs.open sql,conn,1,3
rs("Username")=Username
rs("Password")=Password
rs("Level")=Level
rs.update
rs.close
set rs=nothing
response.write"<Script>alert('�û��޸ĳɹ�');location.href='user.asp';</Script>"
response.end
end if
%>