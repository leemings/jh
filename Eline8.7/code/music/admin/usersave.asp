<!--#include file="function.asp"-->
<%CheckAdmin3%>
<!--#include file="conn.asp"-->
<%
id=request.QueryString("id")
act=request("act")

set rs=server.CreateObject("ADODB.RecordSet")
if act="edit" then
	username=trim(request.form("username"))
	password=trim(request.form("password"))
	Sex=trim(request.form("Sex"))
	Email=trim(request.form("Email"))
	Tel=trim(request.form("Tel"))
	Name=trim(request.form("Name"))
	Address=trim(request.form("Address"))
	Shenfenzheng=trim(request.form("Shenfenzheng"))
	if username="" or password="" or Sex="" then
		errmsg="<br>用户名称，密码和性别都不能为空！</li>"
		call error()
		Response.End 
	end if
	sql="select * from [user] where id="& request("id")
	rs.open sql,conn,1,3
	if not rs.eof then
		rs("Username")=username
		rs("Password")=password
		rs("Sex")=Sex

		if Email="" then
			rs("Email")=null
		else
			rs("Email")=Email
		end if

		if Tel="" then
			rs("Tel")=null
		else
			rs("Tel")=Tel
		end if

		if Name="" then
			rs("Name")=null
		else
			rs("Name")=Name
		end if

		if Address="" then
			rs("Address")=null
		else
			rs("Address")=Address
		end if

		if Shenfenzheng="" then
			rs("Shenfenzheng")=null
		else
			rs("Shenfenzheng")=Shenfenzheng
		end if

		rs.update
	end if
	rs.close
elseif act="lock" then
	sql="select * from [user] where id="&id
	rs.open sql,conn,1,3
	if not rs.eof then
		if rs("lockuser")=false then
			rs("lockuser")=true
		else
			rs("lockuser")=false
		end if
		rs.update
	end if
	rs.close
end if
set rs=nothing
conn.close
set conn=nothing
response.redirect "UserList.asp"
%>


