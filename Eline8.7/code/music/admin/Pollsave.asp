<!--#include file="function.asp"-->
<%CheckAdmin2%>
<!--#include file="conn.asp"-->
<!--#include file="char.inc"-->
<%
ID=request.QueryString("id")
Title=trim(request.form("Title"))
DateAndTime=trim(request.form("DateAndTime"))
act=request("act")
if DateAndTime="" then DateAndTime=now()
Content=trim(request.form("Content"))
founerr=false
if Title="" then
	errmsg="<li>调查主题不能为空</li>"
	founderr=true
end if

if founderr=true then
	call error()
else
	set rs=server.createobject("adodb.recordset")
	if act="edit" then
		sql="select * from Poll where ID="&ID
	elseif act="add" then
		sql="select * from Poll"
	else
		errmsg="<li>操作错误！请联系管理员</li>"
		call error()
		Response.End 
	end if
	rs.open sql,conn,1,3
	if act="add" or act="edit" then
		if act="edit" then
			if rs.eof then
				errmsg="<li>操作错误！请联系管理员</li>"
				call error()
				Response.End 
			end if
		end if
		if act="add" then rs.addnew
		rs("Title")=Title
		for i=1 to 10
			if request("select"&i)<>"" then
				rs("select"&i)=request("select"&i)
				if request("answer"&i)="" then
					rs("answer"&i)=0
				else
					rs("answer"&i)=request("answer"&i)
				end if
				if request("url"&i)="" then
					rs("url"&i)=0
				else
					rs("url"&i)=request("url"&i)
				end if
			end if
		next
		rs("DateAndTime")=DateAndTime
		rs.update
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "Polllist.asp"
end if
%>




