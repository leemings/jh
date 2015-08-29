<!--#include file="conn.asp"-->
<!--#include file="z_dgconn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'---------------------
' write by 绿水青山
'---------------------
dim id,candeletedg
candeletedg=false	
if membername="" or (not founduser) then
	errmsg=errmsg+"<br>"+"<li>您还没有登录论坛，请登录后再来"
	founderr=true
end if	
if request("id")="" or (not isnumeric(request("id"))) then
	errmsg=errmsg+"<br>"+"<li>参数非法，请确认您是从有效连接进入"
	founderr=true
else
	id=clng(request("id"))
end if
'==========判断是否可以删除点歌 开始
if not founderr then
	set rs=connDG.execute("select sender,incept from [media] where id="&id)
	if rs.bof and rs.eof then
		errmsg=errmsg+"<br>"+"<li>参数非法，请确认您是从有效连接进入"
		founderr=true
	elseif rs(0)=membername or rs(1)=membername or master then
		candeletedg=true
	else
		errmsg=errmsg+"<br>"+"<li>您没有删除此点歌的权限"
		founderr=true			
	end if	
end if		
'==========判断是否可以删除点歌 结束	
if founderr or (not candeletedg) then
	stats="点歌管理"
	call nav()
	call head_var(0,0,"点歌控制面板","z_dglistall.asp")
	call dvbbs_error()
	call CloseDB()
	call footer()
else	
	sql="delete * from [media] where id="&id
	connDG.Execute(sql)
	conn.close
	set conn=nothing
	call CloseDB()
	response.redirect Request.ServerVariables("HTTP_REFERER")
end if
%>