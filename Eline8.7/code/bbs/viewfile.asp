<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<%
dim Downid
stats="下载文件"

if Cint(GroupSetting(49))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有浏览本社区展区的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	founderr=true
end if
if request("id")="" then
	Errmsg=Errmsg+"<br>"+"<li>请指定相关文件。"
	founderr=true
elseif not isInteger(request("id")) then
	Errmsg=Errmsg+"<br>"+"<li>非法的文件参数。"
	founderr=true
else
	DownID=request("id")
end if
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
	call footer()
else
	set rs=conn.execute("select * from dv_upfile where F_id="&downid)
	if rs.eof and rs.bof then
		call nav()
		call head_var(2,0,"","")
		Errmsg="<br><li>没有找到相关文件。"
		call dvbbs_error()
		call footer()
	else
		if rs("f_flag")=0 then
			conn.execute("update dv_upfile set F_DownNum=F_DownNum+1 where F_ID="&DownID)
			response.redirect "UpLoadFile/"&rs("F_filename")
		end if
	end if
end if
%>