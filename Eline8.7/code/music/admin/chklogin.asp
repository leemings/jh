<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->  
<%
username=Checkin(trim(Request.form("username")))
password=Checkin(trim(Request.form("password")))
if username="" or password="" then Response.Redirect ("eline_go.asp")

set rs=server.createobject("adodb.recordset")
sql="select * from admin where username='"&username&"'and password='"&password&"'"
rs.open sql,conn,1,3
if not rs.EOF then
	rs("LoginTimes")=rs("LoginTimes")+1
	rs("LoginTime")=now()
	rs("LoginIP")=Request.ServerVariables("REMOTE_ADDR")
	rs.Update

	Session("AdminID")=rs("id")
	Session("IsAdmin")=true
	Session("KEY")=rs("oskey")
        Session.timeout=900
	Response.Redirect ("eline_go2.asp")
else
	errmsg="请输入正确的管理员名字和密码！"
	call error()
	Response.End 
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>


