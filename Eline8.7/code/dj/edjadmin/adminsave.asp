<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<%
founderr=false

username=request.form("username")
password=request.form("password")
oskey=request.form("oskey")
set rs=server.CreateObject("ADODB.RecordSet")
if request("act")="edit" or act="add" then
if username="" or password="" then
	errmsg=errmsg+"<br>"+"<li>����Ա���ƺ����붼����Ϊ�գ�"
	founderr=true
	call error()
	Response.End 
end if
end if
if request("act")="edit" and request.QueryString("id")<>"" then
	id=request("id")
	sql="select * from admin where id="& request.QueryString("id")
	rs.open sql,conn,3,2
	if not rs.eof then
		rs("Username")=username
		rs("Password")=password
		rs.update
	end if
	rs.close
elseif request("act")="add" then
	sql="select * from admin where username='"&username&"'"
	rs.open sql,conn,3,2
	if (rs.eof and rs.bof) then
		rs.addnew
		rs("Username")=UserName
		rs("Password")=Password
		rs.update
	end if
	rs.close
elseif request("act")="del" then
rs.open "delete * from admin where id="&request.QueryString("id"),conn,1
set rst=nothing
response.redirect "AdminMana.asp"
end if
set rs=nothing
conn.close
set conn=nothing
response.redirect "AdminMana.asp"
%>