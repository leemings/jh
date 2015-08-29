<!--#include file="conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="checkadmin.inc"-->
<%
founderr=false
if trim(request.form("WebSiteName"))="" then
	errmsg=errmsg+"<br>"+"<li>网站名称不能为空！"
	founderr=true
end if
if trim(request.form("WebSiteUrl"))="" then
	errmsg=errmsg+"<br>"+"<li>网站地址不能为空！"
	founderr=true
end if

if founderr=true then
	call error()
else
	set rs=server.createobject("adodb.recordset")
	sql="select * from const"
	rs.open sql,conn,1,3
	rs("WebSiteName")=request.form("WebSiteName")
	rs("WebSiteUrl")=request.form("WebSiteUrl")
	rs("WebMaster")=request.form("WebMaster")
	rs("OICQ")=request.form("OICQ")
	rs("EMail")=request.form("EMail")
	rs("LOGO")=request.form("LOGO")
	rs("newreg")=request.form("newreg")
	rs("freeListen")=request.form("freeListen")
	rs("Collect")=request.form("Collect")
	rs("djserver")=request.form("djserver")
	rs("Urgent")=request.form("Urgent")
		rs.update
	rs.close
	set rs=nothing
        conn.close
        set conn=nothing
	Application.Lock
	Application.UnLock
	Response.Redirect "ConstModify.asp"
end if

%>