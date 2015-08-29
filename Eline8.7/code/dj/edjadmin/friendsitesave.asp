<!--#include file="conn.asp"-->
<!--#include file="../function.asp"-->
<!--#include file="checkadmin.inc"-->
<!--#include file="../inc/char.inc"-->
<%
ID=request.QueryString("id")
SiteName=request.form("SiteName")
SiteUrl=request.form("SiteUrl")
LogoUrl=request.form("LogoUrl")
SiteAdmin=request.form("SiteAdmin")
SiteIntro=request.form("SiteIntro")
act=request("act")

founderr=false
if act="add" or act="edit" then
if SiteName="" then
	conn.close
	set conn=nothing
	response.write "<script language=javascript>alert('对不起，请填写你的网站名称！');history.back(1);</script>"
	Response.End
	end if
if SiteUrl="" then
	conn.close
	set conn=nothing
	response.write "<script language=javascript>alert('对不起，请填写你的网站地址！');history.back(1);</script>"
	Response.End
end if

	set rs=server.createobject("adodb.recordset")
	if act="edit" then sql="select * from FriendSite where ID="&ID
	if act="add" then sql="select * from FriendSite"
	rs.open sql,conn,1,3
	if act="edit" then
		if rs.eof then
			errmsg=errmsg+"<br>"+"<li>操作错误！请联系管理员"
			call error()
			Response.End 
		end if
	end if
	if act="add" then
		rs.addnew
	end if
	rs("SiteName")=SiteName
	rs("SiteUrl")=SiteUrl
	if SiteAdmin="" then
		rs("SiteAdmin")=null
	else
		rs("SiteAdmin")=SiteAdmin
	end if
	if LogoUrl="" then
		rs("LogoUrl")=null
	else
		rs("LogoUrl")=LogoUrl
	end if
	if SiteIntro="" then
		rs("SiteIntro")=null
	else
		rs("SiteIntro")=SiteIntro
	end if
	rs.update
	rs.close
	set rs=nothing
elseif act="del" then
	sql="delete from FriendSite where id=" & ID
	conn.execute sql
elseif act="set" then
	set rs=server.createobject("adodb.recordset")
	sql="Select isOK from FriendSite where IsOK=true"
	rs.open sql,conn,1,3
	if not rs.eof then
		do while not rs.eof
			rs("IsOK")=false
			rs.update
		rs.movenext
		loop
	end if
	rs.close

	Checked=replace(request("checked")," ","")
	CheckedNum=Split(Checked,",")
	HowLong=UBound(checkedNum)
	for i=0 to HowLong
		sql="Select IsOK from FriendSite where ID="&CheckedNum(i)
		rs.open sql,conn,1,3
		if not rs.EOF then
			do while not rs.EOF 
				rs("isOK")=true
				rs.update
			rs.MoveNext
			loop
		end if
		rs.close
	next
	set rs=nothing
end if
conn.close
set conn=nothing
if founderr=true then
	call error()
else
	Response.Redirect "FriendSiteMana.asp"
end if
%>